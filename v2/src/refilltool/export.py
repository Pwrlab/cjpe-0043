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
    ("Reference", "Comparison of methods for dividing nighttime sap flow components in Populus tomentosa. "
                  "Chinese Journal of Plant Ecology, 2023.0043"),
    ("summary", "One row per night x tree x method (long format, ready for R/ggplot2)."),
    ("nights", "Night segmentation, quality flags, daytime radiation, rain and nightly means of covariates."),
    ("height_diff", "Height-difference prediction method: A stem refilling (Q_lower - Q_upper, eq. 2/6), "
                    "B canopy refilling and C transpiration (chosen method on the upper sensor)."),
    ("method_evaluation", "Paper 1.5: C1 vs C2 (transpiration at lower vs upper sensor) and A1 = (A1+B1) - B2 "
                          "vs Ah (eq. 5). The best method has C1 ~ C2 and A1 ~ Ah."),
    ("<method>_ratio", "Refilling / night sap flow, nights x trees (v1 sheet '<method>_rate')."),
    ("<method>_fit_sum", "Nocturnal refilling (v1 '<method>_night_sum')."),
    ("<method>_night_sum", "Night sap flow (v1 '<method>_sum')."),
    ("<method>_fitting", "Observed flow and refilling curve per record (optional)."),
    ("night_sum", "Night sap flow total."),
    ("fit_sum", "Nocturnal refilling."),
    ("remainder", "Nocturnal transpiration = night_sum - fit_sum."),
    ("ratio", "fit_sum / night_sum (share of refilling)."),
    ("breakpoint", "Points from sunset: P2 for Re/Et, P1 for Exp, lag for Line."),
    ("breakpoint_source", "auto = Kneedle; fixed = fallback/fixed value; lag = Line lag."),
    ("status", "ok / skipped reason / excluded reason."),
]


def _wide_sheets(res: Result):
    out = {}
    for m in res.config["methods"]:
        for val in ("ratio", "fit_sum", "night_sum"):
            out[f"{m}_{val}"] = res.wide(m, val).reset_index()
    return out


def fitting_table(res: Result, method: str) -> pd.DataFrame:
    """Rows = records of every used night, columns = observed / refilling curve per tree."""
    frames = []
    for nt in res.night_objs:
        if not any((nt.night_id, t, method) in res.fits for t in res.dataset.tree_cols):
            continue
        d = {"night_id": nt.night_id, "time": res.times(nt.night_id)}
        for t in res.dataset.tree_cols:
            d[f"{t}_obs"] = res.observed(nt.night_id, t)
            f = res.fits.get((nt.night_id, t, method))
            d[f"{t}_refill"] = f if f is not None else np.nan
        frames.append(pd.DataFrame(d))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def method_overview(res: Result) -> pd.DataFrame:
    ok = res.summary[res.summary.status == "ok"]
    if ok.empty:
        return pd.DataFrame()
    g = ok.groupby(["method", "tree"])
    out = g.agg(n_nights=("ratio", "size"), ratio_mean=("ratio", "mean"),
                ratio_se=("ratio", lambda v: v.std(ddof=1) / np.sqrt(len(v))),
                refilling_mean=("fit_sum", "mean"), transpiration_mean=("remainder", "mean"),
                night_sum_mean=("night_sum", "mean")).reset_index()
    out.insert(1, "method_label", out.method.map(METHOD_LABELS))
    return out


def to_excel(res: Result, path_or_buf, fitting=False, source_name=""):
    info = pd.DataFrame(README_ROWS, columns=["item", "description"])
    meta = pd.DataFrame([
        ("Input", source_name),
        ("Run at", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Time step (min)", res.dataset.step_minutes),
        ("Unit of sums", res.unit),
        ("Methods", ", ".join(METHOD_LABELS[m] for m in res.config["methods"])),
        ("Mode", "legacy (v1)" if res.config["legacy"] else "paper (v2)"),
        ("Data messages", " | ".join(res.dataset.messages) or "none"),
        ("Config", config_text(res.config)),
    ], columns=["item", "description"])
    with pd.ExcelWriter(path_or_buf, engine="openpyxl") as w:
        pd.concat([meta, info]).to_excel(w, sheet_name="README", index=False)
        method_overview(res).to_excel(w, sheet_name="overview", index=False)
        res.summary.to_excel(w, sheet_name="summary", index=False)
        res.nights.to_excel(w, sheet_name="nights", index=False)
        if len(res.height):
            res.height.to_excel(w, sheet_name="height_diff", index=False)
        if len(res.evaluation):
            res.evaluation.to_excel(w, sheet_name="method_evaluation", index=False)
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
        for name, df in (("summary", res.summary), ("nights", res.nights),
                         ("height_diff", res.height), ("method_evaluation", res.evaluation)):
            if len(df):
                p = outdir / f"{stem}_{name}.csv"
                df.to_csv(p, index=False, encoding="utf-8-sig")
                files.append(p)
    p = outdir / f"{stem}_config.json"
    p.write_text(config_text(res.config), encoding="utf-8")
    files.append(p)
    return files
