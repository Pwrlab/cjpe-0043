"""Run the partitioning methods on every night and tree."""
from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

from . import __version__
from .io import Dataset, daily_radiation
from .methods import DEFAULT_BREAKPOINTS, DEFAULT_PARAMS, partition, resolve_breakpoints, run_method
from .nights import Night, legacy_nights, split_nights

DEFAULT_CONFIG = {
    "methods": ["refilling", "exp", "line", "et"],
    "breakpoints": copy.deepcopy(DEFAULT_BREAKPOINTS),
    "et_min_strategy": True,       # Et: minimum (paper) or mean late-night flow
    "night": {
        "threshold": 5.0,          # W m-2; night = radiation < threshold (paper 1.3)
        "min_run_minutes": 30,     # ignore day/night switches shorter than this
        "drop_partial": True,      # drop incomplete nights at the start/end of the record
        "drop_gap": True,          # drop nights with missing records
        "min_day_rad_MJ": 0.0,     # clear-day filter: daytime radiation of the same date
        "max_rain_mm": 0.0,        # rain filter (needs a rain column): rain from 00:00 of the
                                   # sunset date to the end of the night must be <= this
    },
    "units": {
        "integrate_time": True,    # Q = sum(Fd * dt) (paper eq. 1); False = plain sum (v1)
        "sapwood_area_m2": {},     # {column: As}; with As, Q = sum(Fd * dt * As) in mL
    },
    "height_diff": {
        "pairs": [],               # [{"name": "Tree1", "lower": "T1_1.3m", "upper": "T1_7m"}]
        "method": "line",          # method applied to the upper sensor (paper: Line)
    },
    "exclude_nights": [],          # night_id or sunset date "YYYY-MM-DD"
    "legacy": False,               # reproduce v1 exactly (segmentation, methods, plain sums)
    "params": copy.deepcopy(DEFAULT_PARAMS),  # v1 parameters, used in legacy mode only
}


def merge_config(user: dict | None) -> dict:
    """Merge a user config into the defaults. Accepts the v1 config.json layout
    (method names at the top level) as well."""
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if not user:
        return cfg
    user = copy.deepcopy(user)
    user.pop("version", None)
    for m in DEFAULT_PARAMS:
        if m in user and isinstance(user[m], dict):  # v1 layout
            cfg["params"][m].update(user.pop(m))
    for k, v in user.items():
        if k == "params":
            for m, p in v.items():
                cfg["params"].setdefault(m, {}).update(p)
        elif isinstance(cfg.get(k), dict) and isinstance(v, dict):
            cfg[k].update(v)
        else:
            cfg[k] = v
    return cfg


def load_config(path) -> dict:
    with open(path, encoding="utf-8") as f:
        return merge_config(json.load(f))


@dataclass
class Result:
    summary: pd.DataFrame          # one row per night x tree x method
    nights: pd.DataFrame           # one row per night
    fits: dict                     # (night_id, tree, method) -> refilling curve
    dataset: Dataset
    config: dict
    night_objs: list
    unit: str = ""
    notes: list = field(default_factory=list)
    height: pd.DataFrame = field(default_factory=pd.DataFrame)       # height-difference components
    evaluation: pd.DataFrame = field(default_factory=pd.DataFrame)   # method evaluation (paper 1.5)

    def observed(self, night_id, tree):
        nt = self.night_objs[night_id - 1]
        return self.dataset.df[tree].values[nt.start:nt.end]

    def times(self, night_id):
        nt = self.night_objs[night_id - 1]
        return self.dataset.df[self.dataset.time_col].values[nt.start:nt.end]

    def wide(self, method, value="ratio"):
        s = self.summary[self.summary.method == method]
        w = s.pivot(index="night_id", columns="tree", values=value)[self.dataset.tree_cols]
        starts = s.drop_duplicates("night_id").set_index("night_id")["night_start"]
        w.index = pd.MultiIndex.from_arrays([w.index, starts.reindex(w.index)],
                                            names=["night_id", "night_start"])
        w.columns.name = None
        return w


def _segment(ds: Dataset, cfg) -> list[Night]:
    t, r = ds.df[ds.time_col], ds.df[ds.rad_col].values
    if cfg["legacy"]:
        return legacy_nights(t, r)
    nc = cfg["night"]
    pts = 0
    if ds.step is not None and ds.step.total_seconds() > 0:
        pts = int(round(nc["min_run_minutes"] * 60 / ds.step.total_seconds()))
    return split_nights(t, r, nc["threshold"], pts, ds.step)


def scale_factors(ds: Dataset, cfg):
    """Per-column factor turning a plain sum into a night total, and the unit label."""
    u = cfg["units"]
    if cfg["legacy"] or not u.get("integrate_time", True) or ds.step is None:
        return {t: 1.0 for t in ds.tree_cols}, "sum of records"
    dt = ds.step.total_seconds()
    areas = {k: v for k, v in (u.get("sapwood_area_m2") or {}).items() if v not in (None, "") and float(v) > 0}
    if areas and all(t in areas for t in ds.tree_cols):
        return {t: dt * float(areas[t]) for t in ds.tree_cols}, "mL night-1 (Fd x dt x As)"
    if areas:
        f = {t: dt * float(areas.get(t, 1.0)) for t in ds.tree_cols}
        return f, "mixed: mL for columns with As, mL m-2 for others"
    return {t: dt for t in ds.tree_cols}, "mL m-2 night-1 (Fd x dt)"


def run(ds: Dataset, config: dict | None = None, progress=None) -> Result:
    cfg = merge_config(config)
    nights = _segment(ds, cfg)
    excl = {str(x) for x in cfg.get("exclude_nights", [])}
    drad = daily_radiation(ds)
    factor, unit = scale_factors(ds, cfg)
    step_min = ds.step_minutes or 10.0
    times = ds.df[ds.time_col]

    night_rows, rows, fits = [], [], {}
    total = max(len(nights), 1)
    for k, nt in enumerate(nights):
        date = nt.start_time.date()
        reason = ""
        if not cfg["legacy"]:
            if nt.partial and cfg["night"]["drop_partial"]:
                reason = "partial night"
            elif nt.gap and cfg["night"]["drop_gap"]:
                reason = "time gap"
        prev = drad.loc[date] if date in drad.index else None
        min_rad = float(cfg["night"].get("min_day_rad_MJ", 0) or 0)
        if not reason and min_rad > 0 and (prev is None or not prev["rad_sum_MJ"] >= min_rad):
            reason = "low daytime radiation"
        rain = np.nan
        if ds.rain_col:
            win = (times >= pd.Timestamp(date)) & (times <= nt.end_time)
            rain = float(ds.df.loc[win, ds.rain_col].sum())
            if not reason and not cfg["legacy"] and rain > float(cfg["night"].get("max_rain_mm", 0)):
                reason = "rain"
        if str(nt.night_id) in excl or str(date) in excl:
            reason = "excluded by user"
        nrow = {
            "night_id": nt.night_id, "night_start": nt.start_time, "night_end": nt.end_time,
            "n_points": nt.n, "duration_h": nt.n * step_min / 60,
            "partial": nt.partial, "gap": nt.gap,
            "day_rad_sum_MJ": None if prev is None else prev["rad_sum_MJ"],
            "day_rad_max": None if prev is None else prev["rad_max"],
            "rain_mm": rain, "status": reason or "used",
        }
        for c in ds.cov_cols:  # nightly mean of environmental factors (paper 2.3)
            nrow[f"night_mean_{c}"] = float(np.nanmean(ds.df[c].values[nt.start:nt.end])) \
                if nt.n and not np.all(np.isnan(ds.df[c].values[nt.start:nt.end])) else np.nan
        night_rows.append(nrow)

        for tree in ds.tree_cols:
            y = ds.df[tree].values[nt.start:nt.end].astype(float)
            bps = None
            if not reason and not cfg["legacy"]:
                bps = resolve_breakpoints(y, step_min, cfg["breakpoints"]) if not np.isnan(y).any() else \
                    resolve_breakpoints(np.nan_to_num(y), step_min, {**cfg["breakpoints"], "mode": "fixed"})
            for m in cfg["methods"]:
                if reason:
                    rows.append(_row(nt, tree, m, None, reason))
                    continue
                if cfg["legacy"]:
                    res = run_method(m, y, cfg["params"].get(m), legacy=True)
                else:
                    res = partition(m, y, bps, min_strategy=cfg.get("et_min_strategy", True))
                fits[(nt.night_id, tree, m)] = res.fit
                rows.append(_row(nt, tree, m, res, res.status, ds, factor[tree]))
        if progress:
            progress((k + 1) / total)

    summary = pd.DataFrame(rows)
    if not summary.empty:
        summary["remainder"] = summary["night_sum"] - summary["fit_sum"]
        cols = ["night_id", "night_start", "night_end", "n_points", "tree", "method",
                "breakpoint", "breakpoint_time", "breakpoint_source",
                "night_sum", "fit_sum", "remainder", "ratio", "status"]
        summary = summary[cols]
    res = Result(summary, pd.DataFrame(night_rows), fits, ds, cfg, nights, unit)
    if cfg["height_diff"]["pairs"] and not summary.empty:
        areas = cfg["units"].get("sapwood_area_m2") or {}
        cols = [c for p in cfg["height_diff"]["pairs"] for c in (p["lower"], p["upper"])]
        missing = [c for c in cols if c not in ds.tree_cols]
        no_area = [c for c in cols if c in ds.tree_cols and not areas.get(c)]
        if missing:
            res.notes.append("Height difference: columns not found: " + ", ".join(missing))
        if no_area:
            res.notes.append("Height difference: no sapwood area (As) for " + ", ".join(no_area) +
                             ". Flux densities are compared directly, valid only if As is the same at both "
                             "heights. Enter As to get volumes (mL).")
        res.height, res.evaluation = height_difference(res)
    return res


def _row(nt: Night, tree, m, res, status, ds=None, f=1.0):
    bp_time = pd.NaT
    if res is not None and ds is not None and not np.isnan(res.breakpoint):
        i = nt.start + int(res.breakpoint)
        if i < len(ds.df):
            bp_time = ds.df[ds.time_col].iloc[i]
    return {
        "night_id": nt.night_id, "night_start": nt.start_time, "night_end": nt.end_time,
        "n_points": nt.n, "tree": tree, "method": m,
        "breakpoint": np.nan if res is None else res.breakpoint, "breakpoint_time": bp_time,
        "breakpoint_source": None if res is None else res.extra.get("bp_source", "legacy"),
        "night_sum": np.nan if res is None else res.night_sum * f,
        "fit_sum": np.nan if res is None else res.fit_sum * f,
        "ratio": np.nan if res is None else res.ratio,
        "status": status,
    }


# ------------------------------------------------------------- height difference (paper 1.4.1, 1.4.3, 1.5)
def _pair_table(res: Result, lower, upper, method):
    s = res.summary[res.summary.status == "ok"]
    lo = s[(s.tree == lower) & (s.method == method)].set_index("night_id")
    up = s[(s.tree == upper) & (s.method == method)].set_index("night_id")
    j = lo[["night_start", "night_sum", "fit_sum", "remainder"]].join(
        up[["night_sum", "fit_sum", "remainder"]], lsuffix="_low", rsuffix="_up", how="inner")
    return j


def height_difference(res: Result):
    """Components of the height-difference prediction method and the evaluation
    of each prediction method (C1 vs C2, A1 vs Ah)."""
    hd = res.config["height_diff"]
    comp_rows, eval_rows = [], []
    for p in hd["pairs"]:
        name, lower, upper = p.get("name") or p["lower"], p["lower"], p["upper"]
        if lower not in res.dataset.tree_cols or upper not in res.dataset.tree_cols:
            continue
        # components with the chosen method (paper 1.4.3; eq. 2 and 6)
        j = _pair_table(res, lower, upper, hd["method"])
        for nid, r in j.iterrows():
            q_low, q_up = r.night_sum_low, r.night_sum_up
            a = max(q_low - q_up, 0.0)                        # eq. 2 / eq. 6
            b, c = r.fit_sum_up, r.remainder_up
            tot = a + b + c
            comp_rows.append({"pair": name, "night_id": nid, "night_start": r.night_start,
                              "Q_lower": q_low, "Q_upper": q_up,
                              "A_stem_refilling": a, "B_canopy_refilling": b, "C_transpiration": c,
                              "A_share": a / tot if tot else np.nan, "B_share": b / tot if tot else np.nan,
                              "C_share": c / tot if tot else np.nan, "method": hd["method"]})
        # evaluation of every method (paper 1.5, eq. 5)
        for m in res.config["methods"]:
            j = _pair_table(res, lower, upper, m)
            if len(j) < 3:
                continue
            c1, c2 = j.remainder_low.values, j.remainder_up.values
            a1 = j.fit_sum_low.values - j.fit_sum_up.values      # eq. 5
            ah = np.maximum(j.night_sum_low.values - j.night_sum_up.values, 0)
            eval_rows.append(_eval_row(name, m, c1, c2, a1, ah))
    comp = pd.DataFrame(comp_rows)
    ev = pd.DataFrame(eval_rows)
    if len(hd["pairs"]) > 1 and len(comp):
        # pooled over pairs
        for m in res.config["methods"]:
            parts = []
            for p in hd["pairs"]:
                if p["lower"] in res.dataset.tree_cols and p["upper"] in res.dataset.tree_cols:
                    parts.append(_pair_table(res, p["lower"], p["upper"], m))
            j = pd.concat(parts) if parts else pd.DataFrame()
            if len(j) >= 3:
                eval_rows.append(_eval_row("all pairs", m, j.remainder_low.values, j.remainder_up.values,
                                           j.fit_sum_low.values - j.fit_sum_up.values,
                                           np.maximum(j.night_sum_low.values - j.night_sum_up.values, 0)))
        ev = pd.DataFrame(eval_rows)
    return comp, ev


def _eval_row(name, m, c1, c2, a1, ah):
    se = lambda v: float(np.std(v, ddof=1) / np.sqrt(len(v)))  # noqa: E731
    return {
        "pair": name, "method": m, "n_nights": len(c1),
        "C1_mean": c1.mean(), "C1_se": se(c1), "C2_mean": c2.mean(), "C2_se": se(c2),
        "C1_vs_C2_P_ttest": stats.ttest_ind(c1, c2).pvalue,
        "C1_vs_C2_P_paired": stats.ttest_rel(c1, c2).pvalue,
        "C_bias_pct": (c1.mean() - c2.mean()) / c2.mean() * 100 if c2.mean() else np.nan,
        "A1_mean": a1.mean(), "A1_se": se(a1), "Ah_mean": ah.mean(), "Ah_se": se(ah),
        "A1_vs_Ah_P_ttest": stats.ttest_ind(a1, ah).pvalue,
        "A1_vs_Ah_P_paired": stats.ttest_rel(a1, ah).pvalue,
        "A_bias_pct": (a1.mean() - ah.mean()) / ah.mean() * 100 if ah.mean() else np.nan,
    }


def config_text(cfg) -> str:
    return json.dumps({"version": __version__, **cfg}, ensure_ascii=False, indent=2, default=str)


def save_config(cfg, path):
    Path(path).write_text(config_text(cfg), encoding="utf-8")
