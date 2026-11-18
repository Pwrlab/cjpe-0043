# PyInstaller spec: pyinstaller packaging/refilltool.spec  (run from the v2 folder)
from pathlib import Path
from PyInstaller.utils.hooks import collect_all, copy_metadata, collect_submodules

root = Path(SPECPATH).parent
datas, binaries, hidden = [], [], []
for pkg in ("streamlit", "plotly", "kneed"):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hidden += h
for pkg in ("streamlit", "plotly", "pandas", "numpy", "scipy", "openpyxl", "kneed", "refilltool"):
    try:
        datas += copy_metadata(pkg)
    except Exception:
        pass
hidden += [m for m in collect_submodules("refilltool") if m != "refilltool.app"] + ["openpyxl", "scipy.optimize", "scipy.stats"]
datas += [(str(root / "src" / "refilltool" / "app.py"), "refilltool")]
example = root.parent / "example.xlsx"
if example.exists():
    datas += [(str(example), ".")]

a = Analysis([str(root / "packaging" / "launcher.py")], pathex=[str(root / "src")],
             binaries=binaries, datas=datas, hiddenimports=hidden,
             excludes=["tkinter", "matplotlib", "IPython", "pytest"], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="RefillingTool", console=True,
          icon=str(root.parent / "favicon.ico") if (root.parent / "favicon.ico").exists() else None)
coll = COLLECT(exe, a.binaries, a.datas, name="RefillingTool")
