"""Write results to Excel / CSV."""
from __future__ import annotations

import io as _io
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

from . import __version__
from .methods import METHOD_LABELS
from .pipeline import Result, config_text

README_ROWS = [
    ("Software", f"Refilling Tool v{__version__} (PWRlab, Beijing Forestry University)"),
    ("Reference", "https://www.plant-ecology.com/CN/10.17521/cjpe.2023.0043"),
    ("summary", "One row per night x tree x method (long format, ready for R/ggplot2)."),
    ("nights", "Night segmentation, quality flags and daytime radiation of the same date."),
    ("<method>_ratio", "fit_sum / night_sum, nights x trees (v1 sheet '<method>_rate')."),
    ("<method>_fit_sum", "Nocturnal refilling (v1 '<method>_night_sum')."),
    ("<method>_night_sum", "Sum of observed night sap flow (v1 '<method>_sum')."),
    ("<method>_fitting", "Observed and fitted curves per night (optional)."),
    ("night_sum", "Sum of observed sap flow over the night (same unit as input x points)."),
    ("fit_sum", "Nocturnal refilling: sum of the refilling curve."),
    ("remainder", "Nocturnal transpiration: night_sum - fit_sum."),
    ("breakpoint", "Index (points from sunset) separating early and late night."),
    ("status", "ok / skipped reason / excluded reason."),
]


def _wide_sheets(res: Result):
    out = {}
    for m in res.config["methods"]:
        for val, name in (("ratio", "ratio"), ("fit_sum", "fit_sum"), ("night_sum", "night_sum")):
            w = res.wide(m, val).reset_index()
            out[f"{m}_{name}"] = w
    return out


def fitting_table(res: Result, method: str) -> pd.DataFrame:
    """Rows = time points of every used night, columns = observed / fitted per tree."""
    frames = []
    for nt in res.night_objs:
        if not any((nt.night_id, t, method) in res.fits for t in res.dataset.tree_cols):
            continue
        d = {"night_id": nt.night_id, "time": res.times(nt.night_id)}
        for t in res.dataset.tree_cols:
            d[f"{t}_obs"] = res.observed(nt.night_id, t)
            f = res.fits.get((nt.night_id, t, method))
            d[f"{t}_fit"] = f if f is not None else np.nan
        frames.append(pd.DataFrame(d))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def to_excel(res: Result, path_or_buf, fitting=False, source_name=""):
    info = pd.DataFrame(README_ROWS, columns=["item", "description"])
    meta = pd.DataFrame([
        ("Input", source_name),
        ("Run at", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Time step (min)", res.dataset.step_minutes),
        ("Methods", ", ".join(METHOD_LABELS[m] for m in res.config["methods"])),
        ("Data messages", " | ".join(res.dataset.messages) or "none"),
        ("Config", config_text(res.config)),
    ], columns=["item", "description"])
    with pd.ExcelWriter(path_or_buf, engine="openpyxl") as w:
        pd.concat([meta, info]).to_excel(w, sheet_name="README", index=False)
        res.summary.to_excel(w, sheet_name="summary", index=False)
        res.nights.to_excel(w, sheet_name="nights", index=False)
        for name, df in _wide_sheets(res).items():
            df.to_excel(w, sheet_name=name[:31], index=False)
        if fitting:
            for m in res.config["methods"]:
                ft = fitting_table(res, m)
                if len(ft) < 1_048_000:
                    ft.to_excel(w, sheet_name=f"{m}_fitting"[:31], index=False)


def excel_bytes(res: Result, fitting=False, source_name="") -> bytes:
    buf = _io.BytesIO()
    to_excel(res, buf, fitting=fitting, source_name=source_name)
    return buf.getvalue()


def save_all(res: Result, outdir, stem="refilling_result", fitting=False, csv=True, source_name=""):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    xlsx = outdir / f"{stem}.xlsx"
    to_excel(res, xlsx, fitting=fitting, source_name=source_name)
    files = [xlsx]
    if csv:
        p = outdir / f"{stem}_summary.csv"
        res.summary.to_csv(p, index=False, encoding="utf-8-sig")
        files.append(p)
    p = outdir / f"{stem}_config.json"
    p.write_text(config_text(res.config), encoding="utf-8")
    files.append(p)
    return files
