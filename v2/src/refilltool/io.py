"""Read input tables and detect time / radiation / tree columns."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

TIME_PAT = re.compile(r"^(time|date|datetime|timestamp|时间|日期)", re.I)
RAD_PAT = re.compile(r"(solar|radiation|^rad|^(rs|sr|rg|swin|sw_in|par)\b|辐射|短波)", re.I)


@dataclass
class Dataset:
    df: pd.DataFrame                    # cleaned table (time, radiation, trees)
    time_col: str
    rad_col: str
    tree_cols: list
    step: pd.Timedelta | None
    messages: list = field(default_factory=list)

    @property
    def step_minutes(self):
        return None if self.step is None else self.step.total_seconds() / 60


def read_table(path_or_buf, sheet=0, name: str | None = None) -> pd.DataFrame:
    name = name or str(getattr(path_or_buf, "name", path_or_buf))
    suffix = Path(name).suffix.lower()
    if suffix in (".csv", ".txt", ".dat"):
        try:
            return pd.read_csv(path_or_buf, sep=None, engine="python", encoding="utf-8-sig")
        except UnicodeDecodeError:
            if hasattr(path_or_buf, "seek"):
                path_or_buf.seek(0)
            return pd.read_csv(path_or_buf, sep=None, engine="python", encoding="gbk")
    return pd.read_excel(path_or_buf, sheet_name=sheet)


def detect_columns(df: pd.DataFrame):
    cols = [str(c) for c in df.columns]
    time_col = next((c for c in cols if TIME_PAT.search(c)), None)
    if time_col is None:  # first column that parses as datetime
        for c in cols:
            if pd.api.types.is_datetime64_any_dtype(df[c]):
                time_col = c
                break
    rad_col = next((c for c in cols if c != time_col and RAD_PAT.search(c)), None)
    trees = [c for c in cols if c not in (time_col, rad_col)
             and pd.api.types.is_numeric_dtype(df[c])]
    return time_col, rad_col, trees


def prepare(df: pd.DataFrame, time_col=None, rad_col=None, tree_cols=None,
            regularize=True) -> Dataset:
    """Clean a raw table. With ``regularize`` the series is put on a regular
    time grid so that missing records appear as NaN (and are then flagged)."""
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    t0, r0, tr0 = detect_columns(df)
    time_col = time_col or t0
    rad_col = rad_col or r0
    tree_cols = list(tree_cols) if tree_cols else tr0
    msgs = []
    if time_col is None or rad_col is None:
        raise ValueError("Cannot find the time or solar radiation column; please select them manually.")
    if not tree_cols:
        raise ValueError("No numeric sap flow columns found.")

    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    bad = df[time_col].isna().sum()
    if bad:
        msgs.append(f"Dropped {bad} rows with unreadable time.")
        df = df.dropna(subset=[time_col])
    for c in [rad_col, *tree_cols]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[[time_col, rad_col, *tree_cols]]

    # Logger clocks often drift by milliseconds (e.g. 00:10:00.001): round to seconds
    df[time_col] = df[time_col].dt.round("s")
    if not df[time_col].is_monotonic_increasing:
        df = df.sort_values(time_col)
        msgs.append("Rows were not in time order and have been sorted.")
    dup = df[time_col].duplicated().sum()
    if dup:
        df = df.drop_duplicates(subset=time_col, keep="first")
        msgs.append(f"Removed {dup} duplicated timestamps.")
    df = df.reset_index(drop=True)

    step = None
    if len(df) > 2:
        step = df[time_col].diff().dropna().mode().iloc[0]
        if regularize and step > pd.Timedelta(0):
            grid = pd.date_range(df[time_col].iloc[0], df[time_col].iloc[-1], freq=step)
            missing = len(grid.difference(df[time_col]))
            offgrid = len(df[time_col][~df[time_col].isin(grid)])
            if missing and not offgrid:
                df = df.set_index(time_col).reindex(grid).rename_axis(time_col).reset_index()
                msgs.append(f"{missing} missing records inserted as NaN (step {step}).")
            elif offgrid:
                msgs.append(f"{offgrid} timestamps are off the regular {step} grid; left as is.")
    n_nan_rad = int(df[rad_col].isna().sum())
    if n_nan_rad:
        msgs.append(f"Solar radiation has {n_nan_rad} missing values (filled from the previous record for night detection).")
    for c in tree_cols:
        n_nan = int(df[c].isna().sum())
        if n_nan:
            msgs.append(f"{c}: {n_nan} missing values; affected nights are skipped.")
    return Dataset(df, time_col, rad_col, tree_cols, step, msgs)


def load(path, sheet=0, **kw) -> Dataset:
    return prepare(read_table(path, sheet=sheet), **kw)


def radiation_qc(ds: Dataset, threshold=5.0):
    """Count suspicious radiation values: > threshold between 22:00 and 03:00,
    <= threshold between 10:00 and 14:00 (local time of the logger)."""
    t, r = ds.df[ds.time_col], ds.df[ds.rad_col]
    h = t.dt.hour
    night_high = int(((h >= 22) | (h < 3)).mul(r > threshold).sum())
    noon_low = int(((h >= 10) & (h < 14)).mul(r <= threshold).sum())
    return {"night_high": night_high, "noon_low": noon_low}


def daily_radiation(ds: Dataset) -> pd.DataFrame:
    """Daily radiation sum (MJ m-2, assuming W m-2 input) and maximum."""
    if ds.step is None:
        return pd.DataFrame()
    sec = ds.step.total_seconds()
    g = ds.df.set_index(ds.time_col)[ds.rad_col].resample("D")
    out = pd.DataFrame({"rad_sum_MJ": g.sum() * sec / 1e6, "rad_max": g.max()})
    out.index = out.index.date
    return out


