"""Small preferences kept between runs (settings.json in the app's data folder).

Never anything secret: the password lives in credentials.py's password store.
"""
from __future__ import annotations

import json
import logging

from gradefiller.paths import app_data_dir

log = logging.getLogger("gradefiller")

# Keys
NEVER_ASK_TO_SAVE_PASSWORD = "never_ask_to_save_password"


def _path():
    return app_data_dir() / "settings.json"


def _load() -> dict:
    try:
        data = json.loads(_path().read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def get(key: str, default=None):
    return _load().get(key, default)


def set(key: str, value) -> None:  # noqa: A001  (module-level name, read as settings.set)
    data = _load()
    data[key] = value
    try:
        _path().write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError as e:
        log.info("Couldn't save settings: %s", e)
