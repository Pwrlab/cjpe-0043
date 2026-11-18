"""Tests: synthetic data checks and a regression test against the v1 code."""
import importlib.util
import sys
import types
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

from refilltool import io, methods, nights, pipeline

REPO = Path(__file__).resolve().parents[2]
EXAMPLE = REPO / "example.xlsx"


# ------------------------------------------------------------------ synthetic data
def synthetic(days=4, step_min=10, gap=False):
    t = pd.date_range("2024-06-01 00:00", periods=days * 24 * 60 // step_min, freq=f"{step_min}min")
    h = t.hour + t.minute / 60
    rad = np.clip(900 * np.sin((h - 6) / 12 * np.pi), 0, None)
    night = rad <= 5
    sf = np.where(night, 0.0, rad / 900)
    # night: exponential refilling decay + constant transpiration baseline
    k = 0
    for i in range(len(t)):
        k = k + 1 if night[i] else 0
        if night[i]:
            sf[i] = 0.3 * np.exp(-0.08 * k) + 0.05
    df = pd.DataFrame({"TIMESTAMP": t, "Rs (W/m2)": rad, "T1": sf, "T2": sf * 1.5})
    if gap:
        df = df.drop(index=range(150, 153))
    return df


def test_detect_columns():
    t, r, trees = io.detect_columns(synthetic())
    assert (t, r, trees) == ("TIMESTAMP", "Rs (W/m2)", ["T1", "T2"])


def test_segmentation_flags_partial_and_gap():
    ds = io.prepare(synthetic(gap=True))
    assert ds.step == pd.Timedelta("10min")
    assert any("missing records" in m for m in ds.messages)
    res = pipeline.run(ds)
    st = res.nights.status.tolist()
    assert st[0] == "partial night" and st[-1] == "partial night"
    assert "time gap" in st
    assert (res.nights.status == "used").sum() >= 2


def test_min_run_merges_radiation_spikes():
    rad = np.r_[np.zeros(10), [50, 50], np.zeros(10), np.full(20, 500)]
    assert len(nights._runs(nights.night_mask(rad, 5, 0))) == 4
    m = nights.night_mask(rad, 5, 3)
    assert m[:22].all() and not m[22:].any()


@pytest.mark.parametrize("name", list(methods.FUNCS))
def test_methods_on_synthetic_night(name):
    x = np.arange(70)
    y = 0.3 * np.exp(-0.08 * x) + 0.05
    r = methods.run_method(name, y, None)
    assert r.status == "ok"
    assert 0 < r.ratio <= 1.0 + 1e-9
    assert np.isclose(r.night_sum, y.sum())


def test_method_edge_cases():
    assert methods.refilling(np.zeros(50)).status == "skipped: all zero"
    assert "NaN" in methods.extended_exp_refilling(np.r_[np.ones(20), np.nan]).status
    assert "too short" in methods.refilling(np.ones(10), default_point=30).status


def test_v1_config_layout_accepted():
    cfg = pipeline.merge_config({"refilling": {"default_point": 12, "kneed_process": True}})
    assert cfg["params"]["refilling"] == {"default_point": 12, "kneed_process": True}
    assert cfg["params"]["exp"]["default_point"] == 10


def test_export_roundtrip(tmp_path):
    from refilltool.export import save_all
    res = pipeline.run(io.prepare(synthetic()))
    files = save_all(res, tmp_path, fitting=True)
    x = pd.read_excel(files[0], sheet_name=None)
    assert {"README", "summary", "nights", "exp_ratio", "et_fitting"} <= set(x)
    assert len(x["summary"]) == len(res.summary)


# ------------------------------------------------------------------ regression vs v1
def _load_v1():
    """Import the original v1 modules with a shim for scipy.misc.derivative."""
    def derivative(func, x0, dx=1.0, n=1, args=(), order=3):
        return (func(x0 + dx) - func(x0 - dx)) / (2 * dx)
    shim = types.ModuleType("scipy.misc")
    shim.derivative = derivative
    sys.modules["scipy.misc"] = shim
    sys.path.insert(0, str(REPO))
    try:
        spec = importlib.util.spec_from_file_location("v1_refilling", REPO / "refilling.py")
        v1 = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(v1)
        spec = importlib.util.spec_from_file_location("v1_main", REPO / "main.py")
        main = importlib.util.module_from_spec(spec)
        sys.modules["refilling"] = v1
        spec.loader.exec_module(main)
        return main
    finally:
        sys.path.remove(str(REPO))


@pytest.mark.skipif(not EXAMPLE.exists(), reason="example.xlsx not found")
def test_legacy_mode_matches_v1(monkeypatch):
    main = _load_v1()
    monkeypatch.chdir(REPO)  # v1 reads config.json from the working directory
    app = main.RefillingTool()
    app.input_file = str(EXAMPLE)
    app.methods = {k: True for k in app.methods}
    app.args = types.SimpleNamespace(verbose=False)
    app.read_file()
    app.extract_data()
    app.calculate_data()

    ds = io.load(EXAMPLE, regularize=False)
    res = pipeline.run(ds, {**pipeline.load_config(REPO / "config.json"), "legacy": True})
    for m in ["refilling", "exp", "line", "et"]:
        for v1_key, v2_key in [("rate", "ratio"), ("night_sum", "fit_sum"), ("sum", "night_sum")]:
            ref = np.array([np.asarray(col, dtype=float) for col in app.res_rate[m][v1_key][2:]]).T
            mine = res.wide(m, v2_key).values.astype(float)
            assert ref.shape == mine.shape
            both = ~np.isnan(ref)
            assert np.allclose(ref[both], mine[both], rtol=1e-6, atol=1e-10), (m, v1_key)
            assert np.isnan(mine[np.isnan(ref)]).all() or v1_key == "sum"


def test_et_follows_paper():
    """Paper 1.4.2 (4): T = n * min(flow after P2); refilling = Qn - T."""
    x = np.arange(70)
    y = 0.3 * np.exp(-0.08 * x) + 0.05
    r = methods.run_method("et", y, None)
    base = y[30:].min()
    assert r.breakpoint == 30  # Re breakpoint P2
    assert np.isclose(r.fit_sum, y.sum() - base * len(y))
    assert np.isclose(r.extra["transpiration_sum"], base * len(y))
