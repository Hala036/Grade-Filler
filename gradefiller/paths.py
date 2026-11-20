"""Where the app stores its per-user state (session cookies, etc.)

Kept out of the current working directory / install directory so a packaged
executable works regardless of where it's launched from or double-clicked.
"""
from pathlib import Path
import sys


def app_data_dir() -> Path:
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    elif sys.platform == "win32":
        import os
        base = Path(os.environ.get("APPDATA", Path.home()))
    else:
        base = Path.home() / ".local" / "share"

    d = base / "GradeFiller"
    d.mkdir(parents=True, exist_ok=True)
    return d


def state_path() -> Path:
    return app_data_dir() / "state.json"


def resource_path(relative: str) -> Path:
    """A file shipped with the app (e.g. "assets/icon.ico").

    From source it's relative to the project folder; inside the packed .exe,
    PyInstaller unpacks bundled files to sys._MEIPASS.
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / relative
