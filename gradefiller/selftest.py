"""A quick check that a packaged build works, run as `GradeFiller.exe --self-test`.

Used by the GitHub build after packing: it loads every module, starts
Playwright's driver from inside the bundle and opens the installed browser
without a window. Writes what happened to selftest.txt in the current folder
and returns 0 on success (the .exe has no console to print to).
"""
from __future__ import annotations

import traceback
from pathlib import Path


def run(report: Path = Path("selftest.txt")) -> int:
    lines = []
    try:
        import gradefiller.app  # noqa: F401  (pulls in flet, pandas, openpyxl…)
        from playwright.sync_api import sync_playwright

        from gradefiller.automation import PortalSession

        from gradefiller import updater
        from gradefiller._version import __version__

        lines.append("imports: ok")
        lines.append(f"version: {__version__}")
        lines.append(f"packed exe: {updater.is_packed_exe()} ({updater.running_exe()})")
        from gradefiller.paths import resource_path

        icon = resource_path("assets/icon.ico")
        lines.append(f"window icon: {'ok' if icon.exists() else 'MISSING'} ({icon})")
        if not icon.exists():
            raise FileNotFoundError(icon)
        with sync_playwright() as p:
            lines.append("playwright driver: ok")
            for channel in PortalSession.BROWSER_CHANNELS:
                if channel is None:
                    continue
                try:
                    b = p.chromium.launch(headless=True, channel=channel)
                except Exception as e:  # noqa: BLE001
                    lines.append(f"{channel}: not available ({str(e).splitlines()[0][:100]})")
                    continue
                page = b.new_page()
                page.set_content("<p>ok</p>")
                lines.append(f"{channel}: ok, version {b.version}, page says {page.inner_text('p')!r}")
                b.close()
                break
            else:
                raise RuntimeError("no installed browser could be started")
        lines.append("RESULT: PASS")
        return 0
    except Exception:  # noqa: BLE001
        lines.append(traceback.format_exc())
        lines.append("RESULT: FAIL")
        return 1
    finally:
        report.write_text("\n".join(lines) + "\n", encoding="utf-8")
