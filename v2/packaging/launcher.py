"""Entry point of the stand-alone app (PyInstaller): starts the local
interface and opens the browser. No Python installation needed."""
import os
import sys
from pathlib import Path


def _app_path() -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1] / "src"))
    return str(base / "refilltool" / "app.py")


def main():
    # Imported (not executed as a page) so that PyInstaller bundles every dependency of app.py
    import plotly.graph_objects  # noqa: F401
    import refilltool.cli  # noqa: F401
    import refilltool.export  # noqa: F401
    from streamlit.web import cli as stcli

    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    port = os.environ.get("REFILLTOOL_PORT", "8501")
    print("Refilling Tool v2 - PWRlab, Beijing Forestry University")
    print(f"Opening http://localhost:{port} in your browser. Close this window to quit.")
    # headless avoids Streamlit's first-run e-mail prompt; we open the browser ourselves
    if os.environ.get("REFILLTOOL_NO_BROWSER") != "1":
        import threading
        import webbrowser
        threading.Timer(3.0, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    sys.argv = ["streamlit", "run", _app_path(), "--global.developmentMode=false",
                "--server.address=localhost", f"--server.port={port}",
                "--server.headless=true", "--browser.gatherUsageStats=false"]
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
