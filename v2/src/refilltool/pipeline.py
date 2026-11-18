"""Run the selected methods on every night and tree."""
from __future__ import annotations

import copy
import json
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd

from . import __version__
from .io import Dataset, daily_radiation
from .methods import DEFAULT_PARAMS, run_method
from .nights import Night, legacy_nights, split_nights

DEFAULT_CONFIG = {
    "methods": ["refilling", "exp", "line", "et"],
    "params": copy.deepcopy(DEFAULT_PARAMS),
    "night": {
        "threshold": 5.0,          # W m-2; night = radiation <= threshold
        "min_run_minutes": 30,     # ignore day/night switches shorter than this
        "drop_partial": True,      # drop incomplete nights at the start/end of the record
        "drop_gap": True,          # drop nights with missing records
        "min_day_rad_MJ": 0.0,     # keep nights whose preceding day had >= this radiation (clear-day filter)
    },
    "exclude_nights": [],          # night_id or sunset date "YYYY-MM-DD" to exclude
    "legacy": False,               # reproduce v1 segmentation exactly
}


def merge_config(user: dict | None) -> dict:
    """Merge a user config into the defaults. Accepts the v1 config.json layout
    (method names at the top level) as well."""
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if not user:
        return cfg
    user = copy.deepcopy(user)
    for m in DEFAULT_PARAMS:
        if m in user and isinstance(user[m], dict):  # v1 layout
            cfg["params"][m].update(user.pop(m))
    for k, v in user.items():
        if k == "params":
            for m, p in v.items():
                cfg["params"].setdefault(m, {}).update(p)
        elif k == "night":
            cfg["night"].update(v)
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
    fits: dict                     # (night_id, tree, method) -> fitted array
    dataset: Dataset
    config: dict
    night_objs: list

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


def run(ds: Dataset, config: dict | None = None, progress=None) -> Result:
    cfg = merge_config(config)
    nights = _segment(ds, cfg)
    excl = {str(x) for x in cfg.get("exclude_nights", [])}
    drad = daily_radiation(ds)

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
        if str(nt.night_id) in excl or str(date) in excl:
            reason = "excluded by user"
        night_rows.append({
            "night_id": nt.night_id, "night_start": nt.start_time, "night_end": nt.end_time,
            "n_points": nt.n,
            "duration_h": nt.n * (ds.step.total_seconds() / 3600 if ds.step is not None else np.nan),
            "partial": nt.partial, "gap": nt.gap,
            "day_rad_sum_MJ": None if prev is None else prev["rad_sum_MJ"],
            "day_rad_max": None if prev is None else prev["rad_max"],
            "status": reason or "used",
        })
        for tree in ds.tree_cols:
            y = ds.df[tree].values[nt.start:nt.end].astype(float)
            for m in cfg["methods"]:
                if reason:
                    rows.append(_row(nt, tree, m, None, reason))
                    continue
                res = run_method(m, y, cfg["params"].get(m), legacy=cfg["legacy"],
                                 re_params=cfg["params"].get("refilling"))
                fits[(nt.night_id, tree, m)] = res.fit
                rows.append(_row(nt, tree, m, res, res.status, ds))
        if progress:
            progress((k + 1) / total)

    summary = pd.DataFrame(rows)
    if not summary.empty:
        summary["remainder"] = summary["night_sum"] - summary["fit_sum"]
        cols = ["night_id", "night_start", "night_end", "n_points", "tree", "method",
                "breakpoint", "breakpoint_time", "night_sum", "fit_sum", "remainder", "ratio", "status"]
        summary = summary[cols]
    return Result(summary, pd.DataFrame(night_rows), fits, ds, cfg, nights)


def _row(nt: Night, tree, m, res, status, ds=None):
    bp_time = pd.NaT
    if res is not None and ds is not None and not np.isnan(res.breakpoint):
        i = nt.start + int(res.breakpoint)
        if i < len(ds.df):
            bp_time = ds.df[ds.time_col].iloc[i]
    return {
        "night_id": nt.night_id, "night_start": nt.start_time, "night_end": nt.end_time,
        "n_points": nt.n, "tree": tree, "method": m,
        "breakpoint": np.nan if res is None else res.breakpoint, "breakpoint_time": bp_time,
        "night_sum": np.nan if res is None else res.night_sum,
        "fit_sum": np.nan if res is None else res.fit_sum,
        "ratio": np.nan if res is None else res.ratio,
        "status": status,
    }


def config_text(cfg) -> str:
    return json.dumps({"version": __version__, **cfg}, ensure_ascii=False, indent=2, default=str)


def save_config(cfg, path):
    Path(path).write_text(config_text(cfg), encoding="utf-8")
