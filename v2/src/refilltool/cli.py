"""Command line interface.

Examples
--------
refilltool gui                                   # open the local web interface
refilltool run data.xlsx -o results              # all four methods
refilltool run data.csv -m refilling exp -c my_config.json --fitting
refilltool run data.xlsx --legacy                # reproduce v1 numbers
refilltool run data.xlsx --pair T1:T1_1.3m:T1_7m --area T1_1.3m=0.012 T1_7m=0.008
                                                 # height-difference method + method evaluation
refilltool init-config my_config.json            # write an editable config
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from . import __version__
from .export import save_all
from .io import detect_covariates, detect_pairs, detect_rain, prepare, radiation_qc, read_table
from .methods import METHOD_LABELS
from .pipeline import DEFAULT_CONFIG, load_config, merge_config, save_config

LOGO = r"""   ___  __    __  __    __    _      ___
  / _ \/ / /\ \ \/__\  / /   /_\    / __\
 / /_)/\ \/  \/ / \// / /   //_\\  /__\//
/ ___/  \  /\  / _  \/ /___/  _  \/ \/  \
\/       \/  \/\/ \_/\____/\_/ \_/\_____/   Refilling Tool v""" + __version__


def cmd_run(a):
    cfg = load_config(a.config) if a.config else merge_config(None)
    if a.methods:
        cfg["methods"] = a.methods
    if a.legacy:
        cfg["legacy"] = True
    if a.threshold is not None:
        cfg["night"]["threshold"] = a.threshold
    if a.min_day_rad is not None:
        cfg["night"]["min_day_rad_MJ"] = a.min_day_rad
    if a.max_rain is not None:
        cfg["night"]["max_rain_mm"] = a.max_rain
    if a.breakpoints:
        cfg["breakpoints"]["mode"] = a.breakpoints
    if a.area:
        for item in a.area:
            col, val = item.rsplit("=", 1)
            cfg["units"]["sapwood_area_m2"][col] = float(val)
    if a.pair:
        pairs = []
        for item in a.pair:
            parts = item.split(":")
            if len(parts) == 2:
                parts = [parts[0], *parts]
            pairs.append({"name": parts[0], "lower": parts[1], "upper": parts[2]})
        cfg["height_diff"]["pairs"] = pairs
    raw = read_table(a.input, sheet=a.sheet)
    if not a.pair and not cfg["height_diff"]["pairs"] and not a.no_auto_pair:
        auto = detect_pairs([str(c) for c in raw.columns])
        if auto:
            cfg["height_diff"]["pairs"] = auto
            print("Auto-paired sensors: " + ", ".join(f"{p['name']} ({p['lower']} / {p['upper']})" for p in auto))
    rain = a.rain_col or detect_rain(raw)
    covs = a.covariates if a.covariates is not None else detect_covariates(raw)
    ds = prepare(raw, time_col=a.time_col, rad_col=a.rad_col, tree_cols=a.trees,
                 regularize=not cfg["legacy"], rain_col=rain, cov_cols=covs)
    print(f"Input : {a.input}")
    print(f"Columns: time='{ds.time_col}', radiation='{ds.rad_col}', {len(ds.tree_cols)} trees")
    print(f"Step  : {ds.step_minutes} min")
    if ds.rain_col:
        print(f"Rain  : '{ds.rain_col}' (nights with rain > {cfg['night']['max_rain_mm']} mm are excluded)")
    if ds.cov_cols:
        print(f"Covariates (nightly means): {', '.join(ds.cov_cols)}")
    for m in ds.messages:
        print(f"  ! {m}")
    qc = radiation_qc(ds, cfg["night"]["threshold"])
    if qc["night_high"] or qc["noon_low"]:
        print(f"  ! Radiation check: {qc['night_high']} high values at 22-03 h, "
              f"{qc['noon_low']} low values at 10-14 h")

    from .pipeline import run

    def bar(p):
        n = int(p * 30)
        print(f"\r  [{'#' * n}{'.' * (30 - n)}] {p:5.0%}", end="", flush=True)

    res = run(ds, cfg, progress=None if a.quiet else bar)
    print()
    used = (res.nights.status == "used").sum()
    print(f"Nights: {len(res.nights)} found, {used} used")
    for st, n in res.nights.status.value_counts().items():
        if st != "used":
            print(f"  - {st}: {n}")
    ok = res.summary[res.summary.status == "ok"]
    if len(ok):
        print("Median ratio (fit_sum / night_sum):")
        for m, v in ok.groupby("method")["ratio"].median().items():
            print(f"  {METHOD_LABELS[m]:<30s} {v:.3f}")
    print(f"Unit  : {res.unit}")
    for n in res.notes:
        print(f"  ! {n}")
    if len(res.evaluation):
        print("Method evaluation (paper 1.5; best: C1 ~ C2 and A1 ~ Ah):")
        cols = ["pair", "method", "n_nights", "C_bias_pct", "C1_vs_C2_P_ttest", "A_bias_pct", "A1_vs_Ah_P_ttest"]
        print(res.evaluation[cols].round(3).to_string(index=False))
    files = save_all(res, a.output, stem=a.name or Path(a.input).stem + "_refilling",
                     fitting=a.fitting, source_name=str(a.input))
    print("Saved:")
    for f in files:
        print(f"  {f}")


def cmd_gui(a):
    app = Path(__file__).with_name("app.py")
    # Always headless: avoids Streamlit's interactive first-run e-mail prompt
    cmd = [sys.executable, "-m", "streamlit", "run", str(app), "--server.port", str(a.port),
           "--server.address", "localhost", "--server.headless", "true",
           "--browser.gatherUsageStats", "false"]
    url = f"http://localhost:{a.port}"
    print(f"Starting local interface at {url}  (Ctrl+C or close this window to stop)")
    if not a.headless:
        import threading
        import webbrowser
        threading.Timer(3.0, lambda: webbrowser.open(url)).start()
    raise SystemExit(subprocess.call(cmd))


def cmd_init(a):
    save_config(DEFAULT_CONFIG, a.path)
    print(f"Config written to {a.path}")


def main(argv=None):
    p = argparse.ArgumentParser(prog="refilltool", description="Partition nighttime sap flow (PWRlab).")
    p.add_argument("--version", action="version", version=__version__)
    sub = p.add_subparsers(dest="cmd")

    r = sub.add_parser("run", help="process a data file")
    r.add_argument("input", help="xlsx / xls / csv file")
    r.add_argument("-o", "--output", default="results", help="output folder (default ./results)")
    r.add_argument("-n", "--name", help="output file stem")
    r.add_argument("-m", "--methods", nargs="+", choices=list(METHOD_LABELS), help="methods to run (default all)")
    r.add_argument("-c", "--config", help="JSON config (v1 config.json also accepted)")
    r.add_argument("--sheet", default=0, help="Excel sheet name or index")
    r.add_argument("--time-col")
    r.add_argument("--rad-col")
    r.add_argument("--trees", nargs="+", help="sap flow columns (default: all numeric columns)")
    r.add_argument("--threshold", type=float, help="night radiation threshold, W m-2 (default 5)")
    r.add_argument("--min-day-rad", type=float, help="clear-day filter: min daytime radiation, MJ m-2")
    r.add_argument("--breakpoints", choices=["auto", "fixed"], help="auto = Kneedle P1/P2 (default), fixed")
    r.add_argument("--rain-col", help="rain column (auto-detected by name)")
    r.add_argument("--max-rain", type=float, help="exclude nights with rain > this (mm, default 0)")
    r.add_argument("--covariates", nargs="*", help="columns to average per night (default: auto)")
    r.add_argument("--area", nargs="+", metavar="COL=As", help="sapwood area (m2) per column, gives mL")
    r.add_argument("--pair", nargs="+", metavar="NAME:LOWER:UPPER",
                   help="sensor pairs for the height-difference method")
    r.add_argument("--no-auto-pair", action="store_true", help="do not pair columns like T1_1.3m / T1_7.0m")
    r.add_argument("--fitting", action="store_true", help="also write refilling curves")
    r.add_argument("--legacy", action="store_true", help="reproduce v1 night segmentation exactly")
    r.add_argument("-q", "--quiet", action="store_true")
    r.set_defaults(func=cmd_run)

    g = sub.add_parser("gui", help="open the local web interface")
    g.add_argument("--port", type=int, default=8501)
    g.add_argument("--headless", action="store_true", help="do not open a browser")
    g.set_defaults(func=cmd_gui)

    i = sub.add_parser("init-config", help="write the default config file")
    i.add_argument("path", nargs="?", default="refilling_config.json")
    i.set_defaults(func=cmd_init)

    a = p.parse_args(argv)
    if not getattr(a, "func", None):
        print(LOGO)
        p.print_help()
        return
    if isinstance(getattr(a, "sheet", None), str) and a.sheet.isdigit():
        a.sheet = int(a.sheet)
    if a.cmd == "run" and not a.quiet:
        print(LOGO)
    a.func(a)


if __name__ == "__main__":
    main()
