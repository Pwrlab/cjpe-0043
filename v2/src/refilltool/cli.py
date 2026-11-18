"""Command line interface.

Examples
--------
refilltool gui                                   # open the local web interface
refilltool run data.xlsx -o results              # all four methods
refilltool run data.csv -m refilling exp -c my_config.json --fitting
refilltool run data.xlsx --legacy                # reproduce v1 numbers
refilltool init-config my_config.json            # write an editable config
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from . import __version__
from .export import save_all
from .io import load, radiation_qc
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
    ds = load(a.input, sheet=a.sheet, time_col=a.time_col, rad_col=a.rad_col,
              tree_cols=a.trees, regularize=not cfg["legacy"])
    print(f"Input : {a.input}")
    print(f"Columns: time='{ds.time_col}', radiation='{ds.rad_col}', {len(ds.tree_cols)} trees")
    print(f"Step  : {ds.step_minutes} min")
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
    files = save_all(res, a.output, stem=a.name or Path(a.input).stem + "_refilling",
                     fitting=a.fitting, source_name=str(a.input))
    print("Saved:")
    for f in files:
        print(f"  {f}")


def cmd_gui(a):
    app = Path(__file__).with_name("app.py")
    cmd = [sys.executable, "-m", "streamlit", "run", str(app), "--server.port", str(a.port),
           "--server.address", "localhost",
           "--browser.gatherUsageStats", "false"]
    if a.headless:
        cmd += ["--server.headless", "true"]
    print("Starting local interface at http://localhost:%d  (Ctrl+C to stop)" % a.port)
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
    r.add_argument("--fitting", action="store_true", help="also write fitted curves")
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
