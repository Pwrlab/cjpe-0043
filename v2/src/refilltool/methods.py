"""Partitioning methods for nighttime sap flow.

Ported from v1 ``refilling.py`` (PWRlab, CJPE 2023.0043). The math of each
method is unchanged. Differences from v1 are limited to robustness:

* ``scipy.misc.derivative`` (removed in SciPy >= 1.12) is replaced by the
  analytic derivative of the fitted exponential (difference < 1e-9).
* Short nights, NaN values and fitting failures return a status message
  instead of crashing the whole run.
* Et method follows the paper (section 1.4.2 (4)): it uses the Re breakpoint
  P2 and returns refilling = observed - min(late-night flow). The v1 code
  differed from the paper and is kept only for ``legacy=True``.

Each method returns a :class:`MethodResult`. ``fit`` is the refilling curve; ``fit_sum`` is nocturnal refilling (v1 sheet ``*_night_sum``),
``night_sum`` is the sum of observed night sap flow (v1 ``*_sum``) and
``ratio = fit_sum / night_sum`` (v1 ``*_rate``).
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
from kneed import KneeLocator
from scipy.optimize import curve_fit

METHOD_LABELS = {
    "refilling": "Re (time segment)",
    "exp": "Exp (exponential decay)",
    "line": "Line (linear decay)",
    "et": "Et (extended transpiration)",
}

DEFAULT_PARAMS = {
    "refilling": {"default_point": 30, "kneed_process": False},
    "line": {"default_point": 10, "kneed_process": False},
    "exp": {"default_point": 10, "kneed_process": False},
    "et": {"default_point": 10, "kneed_process": False, "min_strategy": True},  # breakpoint used only in legacy mode
}


@dataclass
class MethodResult:
    fit: np.ndarray
    breakpoint: float = np.nan
    fit_sum: float = np.nan
    night_sum: float = np.nan
    ratio: float = np.nan
    status: str = "ok"
    extra: dict = field(default_factory=dict)


def _fail(y, status, breakpoint=np.nan):
    return MethodResult(fit=np.full(len(y), np.nan), breakpoint=breakpoint,
                        night_sum=float(np.sum(y)) if not np.isnan(y).any() else np.nan,
                        status=status)


def _summarise(y, fit, bp, **extra):
    fs, ns = float(np.sum(fit)), float(np.sum(y))
    return MethodResult(fit=np.asarray(fit, dtype=float), breakpoint=bp, fit_sum=fs,
                        night_sum=ns, ratio=fs / ns if ns != 0 else np.nan, extra=extra)


def find_breakpoint(y, default_point, kneed_process, online):
    """Breakpoint index between early and late night (v1 logic)."""
    if kneed_process:
        x = np.arange(len(y))
        model = KneeLocator(x.tolist(), y.tolist(), curve="convex", direction="decreasing",
                            online=online, interp_method="polynomial")
        knee = model.knee
        return int(default_point if knee is None else knee)
    return int(default_point)


def _precheck(y, default_point, min_len):
    if np.isnan(y).any():
        return "skipped: NaN in night"
    if np.all(y == 0):
        return "skipped: all zero"
    if len(y) <= max(default_point, min_len):
        return f"skipped: night too short ({len(y)} points)"
    return None


def refilling(y, default_point=30, kneed_process=False) -> MethodResult:
    """Re method: early-night flow before the breakpoint is the refilling part."""
    y = np.asarray(y, dtype=float)
    msg = _precheck(y, default_point, 2)
    if msg:
        return _fail(y, msg)
    bp = find_breakpoint(y, default_point, kneed_process, online=True)
    if bp == 0:
        return _fail(y, "skipped: breakpoint = 0", bp)
    fit = y.copy()
    fit[bp:] = 0
    return _summarise(y, fit, bp)


def _fit_exp(y, bp):
    x = np.arange(len(y))
    y0 = y[0]

    def fn_exp(xx, b):
        return y0 * np.exp(-b * xx)

    popt, _ = curve_fit(fn_exp, x[:bp], y[:bp], bounds=(0, [0.5]))
    return float(popt[0])


def extended_line_refilling(y, default_point=10, kneed_process=False) -> MethodResult:
    """Line method: exponential fit before breakpoint, tangent line after it."""
    y = np.asarray(y, dtype=float)
    msg = _precheck(y, default_point, 2)
    if msg:
        return _fail(y, msg)
    bp = find_breakpoint(y, default_point, kneed_process, online=False)
    if bp == 0:
        return _fail(y, "skipped: breakpoint = 0", bp)
    try:
        b = _fit_exp(y, bp)
    except Exception as e:  # noqa: BLE001
        return _fail(y, f"fit failed: {e}", bp)
    x = np.arange(len(y))
    # Analytic slope of y0*exp(-b*x) at the breakpoint (v1: scipy.misc.derivative)
    slope = -b * y[0] * np.exp(-b * bp)
    fit = np.empty_like(y)
    fit[:bp] = y[:bp]
    fit[bp:] = slope * x[bp:] + y[bp] - slope * bp
    fit[fit < 0] = 0
    return _summarise(y, fit, bp, decay_b=b)


def extended_exp_refilling(y, default_point=10, kneed_process=False) -> MethodResult:
    """Exp method: exponential decay from the first night point."""
    y = np.asarray(y, dtype=float)
    msg = _precheck(y, default_point, 2)
    if msg:
        return _fail(y, msg)
    bp = find_breakpoint(y, default_point, kneed_process, online=False)
    if bp == 0:
        return _fail(y, "skipped: breakpoint = 0", bp)
    try:
        b = _fit_exp(y, bp)
    except Exception as e:  # noqa: BLE001
        return _fail(y, f"fit failed: {e}", bp)
    x = np.arange(len(y))
    fit = y[0] * np.exp(-(b * x))
    # v1: steepen the curve until it stays below observations at points 5-14
    step, num = 0.0002, 1
    b_adj = b
    while (y[5:15] - fit[5:15]).min() < 0:
        if num > 500:
            break
        b_adj = b + step * num
        fit = y[0] * np.exp(-(b_adj * x))
        num += 1
    return _summarise(y, fit, bp, decay_b=b, decay_b_adjusted=b_adj)


def extended_transpiration(y, default_point=30, kneed_process=False, min_strategy=True,
                           paper=True, online=True) -> MethodResult:
    """Et method (transpiration inversion), CJPE 2023.0043 section 1.4.2 (4).

    Transpiration rate = minimum sap flow after the Re breakpoint P2, extended
    over the whole night. Transpiration = n * rate; refilling = observed - rate.
    ``fit`` is the refilling curve (observed - rate), consistent with the
    other three methods, so ``fit_sum`` = nocturnal refilling.

    ``paper=False`` reproduces the v1 code, which instead returned
    observed[:bp] + rate (zero after bp) and thus mixed both components.
    """
    y = np.asarray(y, dtype=float)
    msg = _precheck(y, default_point, 2)
    if msg:
        return _fail(y, msg)
    bp = find_breakpoint(y, default_point, kneed_process, online=online)
    if bp == 0:
        return _fail(y, "skipped: breakpoint = 0", bp)
    if paper:
        base = float(y[bp:].min()) if min_strategy else float(np.mean(y[bp:]))
        fit = y - base
        return _summarise(y, fit, bp, transpiration_rate=base,
                          transpiration_sum=base * len(y))
    # ---- v1 code, kept for --legacy
    fit = y.copy()
    fit[default_point:] = 0
    if min_strategy:
        base = float((y[bp:] - fit[bp:]).min())
        fit = fit + base
    else:
        base = float(np.mean(y[bp:]))
        fit = base + 0 * y
    return _summarise(y, fit, bp, baseline=base)


FUNCS = {
    "refilling": refilling,
    "exp": extended_exp_refilling,
    "line": extended_line_refilling,
    "et": extended_transpiration,
}


def run_method(name, y, params, legacy=False, re_params=None) -> MethodResult:
    p = {**DEFAULT_PARAMS[name], **(params or {})}
    kw = dict(default_point=int(p["default_point"]), kneed_process=bool(p["kneed_process"]))
    if name == "et":
        kw["min_strategy"] = bool(p.get("min_strategy", True))
        if legacy:  # v1: own breakpoint (default 10 points, offline knee), v1 curve
            kw.update(paper=False, online=False)
        else:       # paper: use the Re breakpoint P2
            rp = {**DEFAULT_PARAMS["refilling"], **(re_params or {})}
            kw.update(default_point=int(rp["default_point"]), kneed_process=bool(rp["kneed_process"]))
    return FUNCS[name](y, **kw)
