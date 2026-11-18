"""Partitioning methods for nighttime sap flow (CJPE 2023.0043, section 1.4.2).

Two implementations live here:

* ``partition`` (v2 default) follows the paper: Kneedle on smoothed night
  data gives P1 (fast -> slow decay, used by Exp) and P2 (flow stops
  changing, used by Re and Et); Line uses the sunset lag (~90 min);
  Exp is constrained to stay below observed flow over the fitting period
  (exact solution of the paper's gradient-descent step); Et uses the minimum
  late-night flow as the whole-night transpiration rate.
* ``refilling`` / ``extended_*`` are exact ports of the v1 code, used only in
  legacy mode to reproduce v1 numbers. ``scipy.misc.derivative`` (removed in
  SciPy >= 1.12) is replaced by the analytic derivative (difference < 1e-9).

Each method returns a :class:`MethodResult`. ``fit`` is the refilling curve; ``fit_sum`` is nocturnal refilling (v1 sheet ``*_night_sum``),
``night_sum`` is the sum of observed night sap flow (v1 ``*_sum``) and
``ratio = fit_sum / night_sum`` (v1 ``*_rate``).
"""
from __future__ import annotations

import warnings
from dataclasses import dataclass, field

import numpy as np
from kneed import KneeLocator
from scipy.optimize import curve_fit

METHOD_LABELS = {
    "refilling": "Re (time segment)",
    "exp": "Exp (exponential decay)",
    "line": "Line (linear decay)",
    "et": "Et (transpiration inversion)",
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


# =============================================================== v2 (paper)
DEFAULT_BREAKPOINTS = {
    "mode": "auto",            # "auto" = Kneedle (paper 1.5); "fixed" = minutes below
    "smooth_minutes": 30,      # moving-average window before knee detection (paper 1.5)
    "p1_minutes": 100,         # P1 for Exp: fixed value, and fallback when auto fails
    "p2_minutes": 300,         # P2 for Re and Et: fixed value, and fallback
    "line_lag_minutes": 90,    # Line: lag of sap flow behind radiation (paper 1.4.2(2))
    "min_minutes": 30,         # auto knees earlier than this are rejected
}


def smooth(y, window_pts):
    if window_pts <= 1:
        return np.asarray(y, dtype=float)
    k = int(window_pts)
    pad = np.pad(np.asarray(y, dtype=float), (k // 2, k - 1 - k // 2), mode="edge")
    return np.convolve(pad, np.ones(k) / k, mode="valid")


def detect_knees(y, smooth_pts=3):
    """Return (P1, P2) from Kneedle on the smoothed curve, or None where not found.

    P1 = first maximum-curvature point (offline Kneedle);
    P2 = last knee found by online Kneedle, i.e. where the flow stops changing.
    """
    ys = smooth(y, smooth_pts)
    x = np.arange(len(ys))
    kw = dict(curve="convex", direction="decreasing", interp_method="polynomial")
    p1 = p2 = None
    try:
        with warnings.catch_warnings(), np.errstate(all="ignore"):
            warnings.simplefilter("ignore")
            p1 = KneeLocator(x, ys, online=False, **kw).knee
            p2 = KneeLocator(x, ys, online=True, **kw).knee
    except Exception:  # noqa: BLE001
        pass
    p1 = None if p1 is None else int(p1)
    p2 = None if p2 is None else int(p2)
    if p1 is not None and p2 is not None and p2 < p1:
        p1, p2 = p2, p1
    return p1, p2


def resolve_breakpoints(y, step_min, bp_cfg):
    """Breakpoints (in points) for one night: dict with p1, p2, lag and their source."""
    c = {**DEFAULT_BREAKPOINTS, **(bp_cfg or {})}
    pts = lambda minutes: max(1, int(round(minutes / step_min)))  # noqa: E731
    n = len(y)
    out = {"p1": pts(c["p1_minutes"]), "p2": pts(c["p2_minutes"]), "lag": pts(c["line_lag_minutes"]),
           "p1_src": "fixed", "p2_src": "fixed"}
    if c["mode"] == "auto":
        p1, p2 = detect_knees(y, pts(c["smooth_minutes"]))
        lo, hi = pts(c["min_minutes"]), n - 3
        if p1 is not None and lo <= p1 <= hi:
            out.update(p1=p1, p1_src="auto")
        if p2 is not None and lo <= p2 <= hi:
            out.update(p2=p2, p2_src="auto")
        if out["p2"] < out["p1"]:
            out["p2"], out["p2_src"] = out["p1"], out["p1_src"]
    return out


def partition(name, y, bps, min_strategy=True) -> MethodResult:
    """v2 partition of one night. ``fit`` = refilling curve for every method."""
    y = np.asarray(y, dtype=float)
    bp = {"refilling": bps["p2"], "et": bps["p2"], "exp": bps["p1"], "line": bps["lag"]}[name]
    src = {"refilling": bps.get("p2_src"), "et": bps.get("p2_src"),
           "exp": bps.get("p1_src"), "line": "lag"}[name]
    msg = _precheck(y, bp, 3)
    if msg:
        return _fail(y, msg)
    if bp < 2:
        return _fail(y, "skipped: breakpoint < 2 points", bp)
    x = np.arange(len(y))

    if name == "refilling":
        # Re: flow before P2 is refilling, after P2 is transpiration
        fit = y.copy()
        fit[bp:] = 0
        return _summarise(y, fit, bp, bp_source=src)

    if name == "et":
        # Et: min (or mean) late-night flow = whole-night transpiration rate
        base = float(y[bp:].min()) if min_strategy else float(np.mean(y[bp:]))
        # refilling = area between flow and the transpiration rate (>= 0)
        return _summarise(y, np.clip(y - base, 0, None), bp, bp_source=src, transpiration_rate=base)

    if y[0] <= 0:
        return _fail(y, "skipped: sap flow at sunset <= 0", bp)
    try:
        b = _fit_exp(y, bp)
    except Exception as e:  # noqa: BLE001
        return _fail(y, f"fit failed: {e}", bp)

    if name == "exp":
        # Smallest b that keeps y0*exp(-b x) <= observed over the fitting period
        xs, ys = x[1:bp], y[1:bp]
        ok = ys > 0
        b_req = float(np.max(-np.log(ys[ok] / y[0]) / xs[ok])) if ok.any() else b
        b_adj = max(b, b_req)
        fit = np.minimum(y[0] * np.exp(-b_adj * x), y)  # refilling cannot exceed observed flow
        return _summarise(y, fit, bp, bp_source=src, decay_b=b, decay_b_adjusted=b_adj)

    # line: tangent of the fitted decay at the end of the lag period
    slope = -b * y[0] * np.exp(-b * bp)
    fit = np.empty_like(y)
    fit[:bp] = y[:bp]
    fit[bp:] = slope * (x[bp:] - bp) + y[bp]
    fit = np.clip(fit, 0, None)
    fit = np.minimum(fit, y)  # refilling cannot exceed observed flow
    return _summarise(y, fit, bp, bp_source=src, decay_b=b, slope=slope)
