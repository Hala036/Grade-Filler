"""Update notice + self-update for the packed Windows .exe.

At startup the app asks GitHub for the newest Release. If it is newer than the
running version, the app shows "A new version is available. [Update]". Update
then:

  1. downloads the new GradeFiller.exe next to the running one (as *.new),
     checking its size and SHA-256 against what GitHub reports;
  2. renames the running .exe to *.old (Windows can't overwrite a running .exe,
     but it can rename it) and moves the new one into its place;
  3. starts the new .exe and closes this one.

The next start deletes the *.old file. Nothing here touches the UI; app.py
calls these from a background thread.

Only a packed Windows .exe updates itself. Running from source (python main.py)
never checks.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import ssl
import subprocess
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from gradefiller._version import __version__

log = logging.getLogger("gradefiller")

REPO = "Hala036/Grade-Filler"
LATEST_API = f"https://api.github.com/repos/{REPO}/releases/latest"
ASSET_NAME = "GradeFiller.exe"
USER_AGENT = "GradeFiller-updater"


@dataclass
class Release:
    version: str  # "1.0.1"
    url: str  # direct download of GradeFiller.exe
    size: int
    sha256: Optional[str]  # hex, when GitHub reports it


def _open(req: urllib.request.Request, timeout: float):
    """urlopen with the computer's certificates, falling back to certifi's.

    Windows Python trusts the system store (which also knows school/office
    proxies); some other Pythons have no certificates at all.
    """
    if req.type != "https":
        return urllib.request.urlopen(req, timeout=timeout)
    try:
        return urllib.request.urlopen(req, timeout=timeout)
    except urllib.error.URLError as e:
        if not isinstance(e.reason, ssl.SSLCertVerificationError):
            raise
        try:
            import certifi
        except ImportError:
            raise e from None
        ctx = ssl.create_default_context(cafile=certifi.where())
        return urllib.request.urlopen(req, timeout=timeout, context=ctx)


def parse_version(text: str) -> tuple[int, ...]:
    """ "v1.2.3" / "1.2.3" -> (1, 2, 3). Anything odd -> (0,)."""
    parts = []
    for p in text.strip().lstrip("vV").split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        if not digits:
            break
        parts.append(int(digits))
    return tuple(parts) or (0,)


def is_newer(latest: str, current: Optional[str] = None) -> bool:
    return parse_version(latest) > parse_version(current or __version__)


def is_packed_exe() -> bool:
    return bool(getattr(sys, "frozen", False)) and sys.platform == "win32"


def running_exe() -> Path:
    return Path(sys.executable).resolve()


def parse_release(data: dict) -> Optional[Release]:
    """The GitHub "latest release" JSON -> Release, or None if it has no .exe."""
    tag = data.get("tag_name") or ""
    for asset in data.get("assets") or []:
        if asset.get("name") == ASSET_NAME:
            digest = asset.get("digest") or ""
            sha = digest.split(":", 1)[1].lower() if digest.startswith("sha256:") else None
            return Release(
                version=tag.lstrip("vV"),
                url=asset["browser_download_url"],
                size=int(asset.get("size") or 0),
                sha256=sha,
            )
    return None


def check_for_update(timeout: float = 10) -> Optional[Release]:
    """The newest Release if it is newer than this one; None otherwise.

    Raises on network trouble; the caller decides to stay quiet about it.
    """
    req = urllib.request.Request(
        LATEST_API,
        headers={"Accept": "application/vnd.github+json", "User-Agent": USER_AGENT},
    )
    with _open(req, timeout) as resp:
        data = json.load(resp)
    release = parse_release(data)
    if release is None:
        log.info("Update check: latest release %s has no %s", data.get("tag_name"), ASSET_NAME)
        return None
    log.info("Update check: running %s, latest %s", __version__, release.version)
    return release if is_newer(release.version) else None


def can_replace(exe: Path) -> bool:
    """True if we may write next to the .exe (not e.g. a read-only folder)."""
    probe = exe.with_name(exe.name + ".write-test")
    try:
        probe.write_bytes(b"")
        probe.unlink()
        return True
    except OSError:
        return False


def download(
    release: Release,
    dest: Path,
    on_progress: Optional[Callable[[int, int], None]] = None,
    timeout: float = 30,
) -> Path:
    """Download release.url to dest, checked. Raises ValueError if it's not right."""
    tmp = dest.with_name(dest.name + ".part")
    sha = hashlib.sha256()
    done = 0
    req = urllib.request.Request(release.url, headers={"User-Agent": USER_AGENT})
    try:
        with _open(req, timeout) as resp, open(tmp, "wb") as out:
            total = release.size or int(resp.headers.get("Content-Length") or 0)
            while True:
                chunk = resp.read(256 * 1024)
                if not chunk:
                    break
                out.write(chunk)
                sha.update(chunk)
                done += len(chunk)
                if on_progress:
                    on_progress(done, total)
        with open(tmp, "rb") as f:
            if f.read(2) != b"MZ":
                raise ValueError("the downloaded file isn't a Windows program")
        if release.size and done != release.size:
            raise ValueError(f"the download stopped early ({done} of {release.size} bytes)")
        if release.sha256 and sha.hexdigest() != release.sha256:
            raise ValueError("the downloaded file is damaged (checksum doesn't match)")
        os.replace(tmp, dest)
        return dest
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


def swap_in(new_file: Path, exe: Path) -> Path:
    """Put new_file where exe is. Returns where the old .exe went (*.old).

    On failure, puts the old .exe back and re-raises.
    """
    old = exe.with_name(exe.name + ".old")
    if old.exists():
        old.unlink()  # left over from an earlier update
    os.replace(exe, old)  # allowed on Windows even while exe is running
    try:
        os.replace(new_file, exe)
    except OSError:
        os.replace(old, exe)
        raise
    return old


def restart(exe: Path) -> None:
    """Start the (new) exe as a separate program."""
    env = dict(os.environ)
    # Tell the new PyInstaller program to unpack itself afresh instead of
    # reusing this one's temporary folder (which goes away when we close).
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    for k in list(env):
        if k.startswith("_PYI_") or k == "_MEIPASS2":
            env.pop(k)
    flags = 0
    if sys.platform == "win32":
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    subprocess.Popen([str(exe)], env=env, cwd=str(exe.parent), close_fds=True, creationflags=flags)


def install_update(release: Release, on_progress: Optional[Callable[[int, int], None]] = None) -> None:
    """Download, swap and start the new version. The caller then closes the app."""
    exe = running_exe()
    new_file = exe.with_name(exe.name + ".new")
    download(release, new_file, on_progress)
    try:
        swap_in(new_file, exe)
    finally:
        if new_file.exists():
            new_file.unlink()
    log.info("Updated %s to %s; restarting", exe, release.version)
    restart(exe)


def clean_up_old(tries: int = 15, wait: float = 2.0) -> None:
    """Delete the *.old (and stray *.new / *.part) left by the last update.

    Right after an update the old program is still closing its browser, and
    Windows won't delete a file that is still running, so keep trying for up
    to ~30 s. Whatever is left is tried again at the next start.
    """
    if not is_packed_exe():
        return
    exe = running_exe()
    left = [exe.with_name(exe.name + s) for s in (".old", ".new", ".new.part")]
    for attempt in range(tries):
        if attempt:
            time.sleep(wait)
        still = []
        for leftover in left:
            try:
                leftover.unlink()
                log.info("Removed %s", leftover)
            except FileNotFoundError:
                pass
            except OSError as e:
                if attempt == tries - 1:
                    log.info("Couldn't remove %s yet: %s", leftover, e)
                still.append(leftover)
        left = still
        if not left:
            return
