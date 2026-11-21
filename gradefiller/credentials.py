"""The saved portal login (username + password), if the teacher chose to save it.

Kept in the computer's own password store via `keyring`: Windows Credential
Manager on Windows (encrypted by Windows, tied to the Windows account), the
Keychain on macOS. Never in a plain file, never in the log.

One saved account at a time. Every function here fails quietly: if the
password store can't be used, the app just doesn't offer to save.
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from typing import Optional

log = logging.getLogger("gradefiller")

SERVICE = "Grade Filler (portal.alquds.edu)"
ENTRY = "saved-login"


@dataclass(frozen=True)
class Login:
    username: str
    password: str


def _keyring():
    import keyring

    kr = keyring.get_keyring()
    name = type(kr).__module__ + "." + type(kr).__name__
    if "fail" in name.lower() or "null" in name.lower():
        raise RuntimeError(f"no password store ({name})")
    return keyring


def available() -> bool:
    try:
        _keyring()
        return True
    except Exception as e:  # noqa: BLE001
        log.info("Password store not available: %s", e)
        return False


def load() -> Optional[Login]:
    try:
        raw = _keyring().get_password(SERVICE, ENTRY)
        if not raw:
            return None
        data = json.loads(raw)
        if data.get("username") and data.get("password"):
            return Login(data["username"], data["password"])
    except Exception as e:  # noqa: BLE001
        log.info("Couldn't read the saved login: %s", type(e).__name__)
    return None


def save(login: Login) -> bool:
    try:
        _keyring().set_password(
            SERVICE, ENTRY, json.dumps({"username": login.username, "password": login.password})
        )
        log.info("Saved the portal login")
        return True
    except Exception as e:  # noqa: BLE001
        log.info("Couldn't save the login: %s", type(e).__name__)
        return False


def forget() -> None:
    try:
        kr = _keyring()
        if kr.get_password(SERVICE, ENTRY) is not None:
            kr.delete_password(SERVICE, ENTRY)
            log.info("Forgot the saved portal login")
    except Exception as e:  # noqa: BLE001
        log.info("Couldn't remove the saved login: %s", type(e).__name__)
