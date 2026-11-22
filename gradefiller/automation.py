"""Playwright-driven automation against the Al-Quds grades portal.

All portal/browser interaction lives here so the UI layer never touches
Playwright directly. Nothing in this module is Tkinter- or Flet-specific.

Nothing here knows about specific courses or exams: the grid's columns are
discovered at runtime and the caller supplies any number of
Excel-column -> portal-column mappings.
"""
from __future__ import annotations

import json
import logging
import math
import re
import sys
import time
from dataclasses import dataclass, field
from decimal import ROUND_FLOOR, ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from typing import Callable, Optional

import openpyxl
import pandas as pd
from playwright.sync_api import (
    sync_playwright,
    Browser,
    BrowserContext,
    Frame,
    Page,
    TimeoutError as PwTimeoutError,
    Error as PwError,
)

COURSE_PAGE_URL = "https://portal.alquds.edu/en/registration/marks/course_list"
# The university's single sign-on page the portal sends you to for logging in
# (a SimpleSAMLphp form: #username, #password, #submit_button).
SSO_HOST = "login.alquds.edu"

# Runs in every page of our browser, before the page's own scripts; acts only
# on the login page. It (1) tells Python what the teacher typed when they log
# in, so the app can offer to save it, and (2) if a login is saved, fills it in
# and presses Log in. It presses at most once per 2 minutes per tab, so a wrong
# saved password never loops (the teacher then sees the page's error and types
# the new one). The saved login comes from window.__gfLogin, set by
# _LOGIN_CONFIG_JS; the newest of those wins, because this waits until the page
# has loaded.
_LOGIN_PAGE_JS = r"""
(() => {
  const host = %s;
  if (location.hostname !== host) return;
  const ready = () => {
    const u = document.querySelector('#username'), p = document.querySelector('#password');
    if (!u || !p) return;
    const report = () => { try { window.gfLoginTyped(u.value, p.value); } catch (e) {} };
    const form = p.form;
    if (form) form.addEventListener('submit', report, true);
    document.addEventListener('click', e => {
      if (e.target.closest && e.target.closest('#submit_button, [type=submit]')) report();
    }, true);
    p.addEventListener('keydown', e => { if (e.key === 'Enter') report(); }, true);

    const saved = window.__gfLogin;
    if (!saved) return;
    if (u.value || p.value) {  // the page kept the username: the login just failed
      window.__gfLoginFailed = true;
      return;
    }
    u.value = saved.username;
    p.value = saved.password;
    let last = 0;
    try { last = +sessionStorage.getItem('gfAutoLogin') || 0; } catch (e) {}
    if (Date.now() - last < 120000) return;
    try { sessionStorage.setItem('gfAutoLogin', String(Date.now())); } catch (e) {}
    const btn = document.querySelector('#submit_button') || (form && form.querySelector('[type=submit], button'));
    if (form && form.requestSubmit) form.requestSubmit(btn && btn.form === form ? btn : undefined);
    else if (btn) btn.click();
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', ready);
  else ready();
})();
"""
_LOGIN_CONFIG_JS = "if (location.hostname === %s) window.__gfLogin = %s;"

log = logging.getLogger("gradefiller")


def _ms(since: float) -> float:
    return (time.perf_counter() - since) * 1000


def _short_url(frame) -> str:
    try:
        url = frame.url or ""
    except PwError:
        return ""
    url = url.split("?", 1)[0]  # drop query strings (may carry session data)
    return url if len(url) <= 70 else url[:67] + "..."


# DHTMLX grid row selector. Rows come in two alternating classes.
ROW_LOCATOR = "tr.ev_dhx_web, tr.odd_dhx_web"

PACE_SECONDS_DEFAULT = 0.12


# ---------- value normalization ----------

def _is_blank(x) -> bool:
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except (TypeError, ValueError):
        pass
    return str(x).strip() == ""


def normalize_id(x) -> str:
    """Turn an ID cell (int, float like 22010083.0, or text) into a clean string.

    Returns "" for blanks. Non-numeric IDs are kept as trimmed text rather
    than crashing.
    """
    if _is_blank(x):
        return ""
    s = str(x).strip()
    try:
        f = float(s)
    except ValueError:
        return s
    if math.isfinite(f) and f.is_integer():
        return str(int(f))
    return s


def parse_grade(x) -> Optional[str]:
    """The text to type for an Excel grade cell, unchanged (no rounding).

    Cells come from load_sheet already as the text Excel displays, so a
    grade shown as 12.5 is typed as 12.5. None for blank; ValueError for
    anything that isn't a number.
    """
    if _is_blank(x):
        return None
    s = str(x).strip()
    f = float(s)
    if not math.isfinite(f):
        raise ValueError(f"not a number: {x!r}")
    return s


ROUND_STEPS = 20  # the round-up line goes 0..1 in steps of 0.05


def step_value(n: int) -> Decimal:
    """A point on the round-up line (0..ROUND_STEPS) as a number: 3 -> 0.15."""
    return Decimal(n) / ROUND_STEPS


def step_label(n: int) -> str:
    """3 -> "0.15", 10 -> "0.5", 20 -> "1"."""
    return format(step_value(n).normalize(), "f")


@dataclass(frozen=True)
class RoundUp:
    """Optional rounding, decided by the part after the decimal point.

    Two joined arrows on a 0-1 line, in steps of 0.05 (0..20),
    start <= joint <= end:
      * below start: the decimals are dropped (12.1 -> 12)
      * start up to joint: up to .joint   (the first, amber arrow)
      * joint up to end:   up to .end     (the second, green arrow;
        end 20 means the next whole number)
      * end and above stays, and whole numbers never change.
    The teacher's usual rule is (2, 10, 20): .1-.49 -> .5, .5-.99 -> next
    whole, and .01-.09 -> .0. Only the decimals below the first arrow ever go
    down (the user's choice, 2026-11-22).
    """
    start: int = 2
    joint: int = 10
    end: int = 20

    def __post_init__(self):
        if not 0 <= self.start <= self.joint <= self.end <= ROUND_STEPS:
            raise ValueError(f"round-up points out of order: {self}")

    def target(self, frac: Decimal) -> Optional[Decimal]:
        """Where a decimal part (0 < frac < 1) goes; None if it stays."""
        if frac < step_value(self.start):
            return Decimal(0)
        if frac < step_value(self.joint):
            return step_value(self.joint)
        if frac < step_value(self.end):
            return step_value(self.end)
        return None

    @property
    def active(self) -> bool:
        return self.start < self.joint or self.joint < self.end


def apply_round_up(text: str, rule: RoundUp) -> str:
    """Round a grade's text by the rule; returns the text to type."""
    try:
        d = Decimal(text.strip())
    except InvalidOperation:
        return text
    if not d.is_finite() or d < 0:
        return text
    whole = d.to_integral_value(rounding=ROUND_FLOOR)
    frac = d - whole
    if frac == 0:
        return text
    target = rule.target(frac)
    if target is None:
        return text
    return format((whole + target).normalize(), "f")


def looks_numeric(values, threshold: float = 0.6) -> bool:
    """True if most non-blank values parse as numbers."""
    non_blank = [v for v in values if not _is_blank(v)]
    if not non_blank:
        return False
    ok = 0
    for v in non_blank:
        try:
            float(str(v).strip())
            ok += 1
        except ValueError:
            pass
    return ok / len(non_blank) >= threshold


# ---------- data classes ----------

@dataclass
class ColumnMapping:
    """One portal grid column (by 0-based td index) mapped to an Excel column."""
    portal_index: int
    portal_label: str
    excel_column: str


@dataclass
class FillJob:
    excel_path: str
    sheet_name: str
    id_mapping: ColumnMapping
    grade_mappings: list[ColumnMapping] = field(default_factory=list)
    zero_as_blank: bool = True  # a 0 in Excel leaves the portal cell as it is, like a blank cell
    round_up: Optional[RoundUp] = None  # None: type grades exactly as Excel shows them
    pace_seconds: float = PACE_SECONDS_DEFAULT


@dataclass
class GridColumn:
    """A column discovered by scanning the live portal grid."""
    index: int
    label: str
    values: list[str] = field(default_factory=list)  # one per grid row

    @property
    def sample_value(self) -> str:
        return next((v for v in self.values if v), "")


@dataclass
class GridScan:
    columns: list[GridColumn]
    labels_found: bool  # False -> labels are generic "Column N"


# ---------- browser session ----------

class PortalSession:
    """Owns one Playwright browser for the whole wizard flow.

    The same browser window is reused for login, navigation and filling.
    If the user closes it, the next call transparently opens a new one
    (with the saved login, if any).
    """

    def __init__(self, state_path: Path, saved_login=None, sso_host: str = SSO_HOST):
        self.state_path = state_path
        self.sso_host = sso_host
        self.saved_login = saved_login  # credentials.Login or None: filled in for the teacher
        # What the teacher last typed on the login page (username, password),
        # kept in memory only, until the app asks whether to save it.
        self.typed_login: Optional[tuple[str, str]] = None
        self.stop_waiting = False  # set from the UI to end wait_for_manual_login early
        self._playwright = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self.grid: Optional[GridInfo] = None  # the table chosen for filling

    def _alive(self) -> bool:
        try:
            return bool(self.browser and self.browser.is_connected() and self.context)
        except PwError:
            return False

    # Browsers to try, in order: the Edge every Windows laptop has, then
    # Chrome, then Playwright's own Chromium (only on dev machines that ran
    # "playwright install chromium"). Using an installed browser keeps the
    # .exe small and needs no download on the teacher's laptop.
    BROWSER_CHANNELS = ("msedge", "chrome", None)
    HEADLESS = False  # tests run without a window

    def _launch_browser(self) -> Browser:
        errors = []
        for channel in self.BROWSER_CHANNELS:
            try:
                browser = self._playwright.chromium.launch(headless=self.HEADLESS, channel=channel)
                log.info("browser: %s", channel or "playwright chromium")
                return browser
            except PwError as e:
                errors.append(f"{channel or 'chromium'}: {str(e).splitlines()[0][:120]}")
        log.info("no browser could be started: %s", "; ".join(errors))
        raise RuntimeError(
            "Couldn't open a web browser. Please install Microsoft Edge or Google Chrome and try again."
        )

    def _ensure_browser(self) -> None:
        if self._alive():
            if self.page is None or self.page.is_closed():
                pages = [p for p in self.context.pages if not p.is_closed()]
                self.page = pages[-1] if pages else self.context.new_page()
            return

        self.close()
        self._playwright = sync_playwright().start()
        self.browser = self._launch_browser()
        self.context = None
        if self.state_path.exists():
            try:
                self.context = self.browser.new_context(storage_state=str(self.state_path))
            except (PwError, ValueError, OSError):
                # Corrupt/old session file: start fresh, user just logs in again.
                self.context = None
        if self.context is None:
            self.context = self.browser.new_context()
        self._watch_login_page()
        self.page = self.context.new_page()

    def _watch_login_page(self) -> None:
        def typed(_source, username, password):
            if username and password:
                login = (str(username).strip(), str(password))
                if login != self.typed_login:
                    log.info("Login page: a username and password were entered")
                self.typed_login = login

        self.context.expose_binding("gfLoginTyped", typed)
        self.context.add_init_script(_LOGIN_PAGE_JS % json.dumps(self.sso_host))
        self._send_saved_login()

    def _send_saved_login(self) -> None:
        login = self.saved_login
        data = {"username": login.username, "password": login.password} if login else None
        self.context.add_init_script(_LOGIN_CONFIG_JS % (json.dumps(self.sso_host), json.dumps(data)))

    def set_saved_login(self, login) -> None:
        """Use this login (or None) on the next login page, without restarting the browser."""
        self.saved_login = login
        if self._alive():
            self._send_saved_login()

    def take_typed_login(self) -> Optional[tuple[str, str]]:
        """What the teacher typed on the login page, once they're past it (then
        cleared). None while still on the login page: it may have been wrong."""
        if self.typed_login is None or self.on_login_page():
            return None
        typed, self.typed_login = self.typed_login, None
        return typed

    def on_login_page(self) -> bool:
        try:
            return bool(self.page and not self.page.is_closed() and self.sso_host in self.page.url)
        except PwError:
            return False

    def log_out(self) -> None:
        """Forget the portal login in the browser and on disk (for "Switch account")."""
        self.typed_login = None
        if self._alive():
            self.context.clear_cookies()
        try:
            self.state_path.unlink()
        except FileNotFoundError:
            pass

    def open_course_list(self, url: str = COURSE_PAGE_URL) -> None:
        """Open the course list. If the portal asks to log in and a login is
        saved, the login page signs in by itself; wait for that and come back."""
        self._ensure_browser()
        self.page.goto(url, wait_until="domcontentloaded")
        self.page.bring_to_front()
        if self.saved_login and self.on_login_page():
            if not self._wait_for_auto_login():
                log.info("Saved login didn't get past the login page")
                return  # wrong/old password: the teacher logs in by hand
            log.info("Logged in with the saved password")
            self.save_session()  # keep the fresh login
            if self.page.url.split("?")[0] != url:
                self.page.goto(url, wait_until="domcontentloaded")

    def wait_for_manual_login(self, url: str = COURSE_PAGE_URL, timeout_s: float = 600) -> bool:
        """While the teacher logs in by hand: wait until they're past the login
        page, keep the new login and go to url. False if they closed the
        browser, took too long, or stop_waiting was set."""
        self.stop_waiting = False
        deadline = time.monotonic() + timeout_s
        while time.monotonic() < deadline and not self.stop_waiting:
            try:
                if not self._alive() or self.page is None or self.page.is_closed():
                    return False
                if self.sso_host not in self.page.url:
                    self.page.wait_for_load_state("domcontentloaded")
                    log.info("Logged in by hand")
                    self.save_session()
                    if self.page.url.split("?")[0] != url:
                        self.page.goto(url, wait_until="domcontentloaded")
                    return True
                self.page.wait_for_timeout(500)
            except PwError:
                return False
        return False

    def _wait_for_auto_login(self, timeout_s: float = 20) -> bool:
        """True once the browser has left the login page; False if the login
        page said no (or nothing happened in time)."""
        deadline = time.monotonic() + timeout_s
        while time.monotonic() < deadline:
            if not self.on_login_page():
                try:
                    self.page.wait_for_load_state("domcontentloaded")
                except PwError:
                    pass
                return True
            try:
                if self.page.evaluate("() => window.__gfLoginFailed === true"):
                    return False
            except PwError:
                pass  # mid-navigation
            self.page.wait_for_timeout(250)
        return False

    def save_session(self) -> None:
        if not self._alive():
            raise RuntimeError("The browser window was closed. Open it again and log in.")
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        self.context.storage_state(path=str(self.state_path))

    def find_grids(self, timeout_ms: int = 5000) -> list["GridInfo"]:
        """Every grade-style table in every tab and frame, best guess first.

        The portal shows a list of groups as one table and opens the real
        grade table in a popup on top of it, so several tables are normal.
        The one in front (popup / newest window) is ranked first.
        """
        if not self._alive():
            return []
        started = time.perf_counter()
        deadline = time.monotonic() + timeout_ms / 1000
        attempt = 0
        while True:
            attempt += 1
            grids: list[GridInfo] = []
            for pi, p in enumerate(self.context.pages):
                if p.is_closed():
                    continue
                for fi, f in enumerate(p.frames):
                    t0 = time.perf_counter()
                    try:
                        found = f.evaluate(_DISCOVER_JS, [ROW_LOCATOR, GRID_ATTR, f"p{pi}f{fi}g", POPUP_TITLE])
                    except PwError as e:
                        log.info("  tab %d frame %d: skipped after %.0f ms (%s) %s", pi, fi,
                                 _ms(t0), str(e).splitlines()[0][:60], _short_url(f))
                        continue  # frame navigated/detached mid-scan
                    in_frame = f != p.main_frame
                    if in_frame and found and not any(d["titled"] for d in found):
                        # The popup's title may sit outside the frame, above it.
                        framed_title = _frame_is_titled(f)
                        if framed_title and len(found) == 1:
                            found[0]["titled"] = True
                    log.info("  tab %d frame %d: %.0f ms, %d table(s)%s %s", pi, fi, _ms(t0), len(found),
                             " [Mark Entry]" if any(d["titled"] for d in found) else "", _short_url(f))
                    for d in found:
                        grids.append(GridInfo(page=p, frame=f, page_index=pi, in_frame=in_frame, **d))
            if grids or time.monotonic() >= deadline:
                grids.sort(key=GridInfo.priority, reverse=True)
                log.info("find_grids: %d table(s) in %.0f ms (%d attempt(s))", len(grids), _ms(started), attempt)
                return grids
            time.sleep(0.25)

    def close(self) -> None:
        try:
            if self.browser:
                self.browser.close()
        except PwError:
            pass
        finally:
            try:
                if self._playwright:
                    self._playwright.stop()
            except PwError:
                pass
            self._playwright = None
            self.browser = None
            self.context = None
            self.page = None
            self.grid = None


# ---------- grid discovery & scanning ----------

# The row selector (tr.ev_dhx_web / tr.odd_dhx_web) is the classic dhtmlxGrid
# skin's row class naming. A page can hold several such grids (the portal's
# group list, plus the grade table in a popup on top of it), so rows are
# grouped by the <table> they belong to, and each grid's container is tagged
# with GRID_ATTR so later steps only ever touch that one grid. dhtmlxGrid
# renders the header as a separate table in the same container, roughly:
#   div.gridbox
#     div.xhdr > table.hdr > tr (sizing row) + tr > td > div.hdrcell
#     div.objbox > table.obj > tr.ev_dhx_web/.odd_dhx_web > td
GRID_ATTR = "data-gradefiller-grid"

# The portal opens the table to fill in a popup with this title. A table
# under it is chosen without asking (lower-case; matched as a substring).
POPUP_TITLE = "mark entry"


_TITLE_HELPER_JS = r"""
function gfTitled(start, attr, needle) {
  // True if `start` sits inside a floating popup (absolutely/fixed positioned
  // box) whose text includes the title. Requiring the popup box keeps a page
  // menu link with the same words from marking the table behind the popup.
  const norm = s => (s || '').toLowerCase().replace(/\s+/g, ' ');
  const floating = el => ['absolute', 'fixed'].includes(getComputedStyle(el).position);
  let inPopup = floating(start);
  for (let e = start.parentElement; e && e !== document.documentElement; e = e.parentElement) {
    // Stop before a level that also holds another table: its title isn't ours.
    const others = Array.from(e.querySelectorAll('[' + attr + ']'))
      .some(o => o !== start && !start.contains(o) && !o.contains(start));
    if (others) return false;
    inPopup = inPopup || floating(e);
    if (inPopup && norm(e.innerText).includes(needle)) return true;
  }
  return false;
}
"""

_DISCOVER_JS = "(" + r"""
([rowSel, attr, prefix, needle]) => {
  """ + _TITLE_HELPER_JS + r"""
  document.querySelectorAll('[' + attr + ']').forEach(e => e.removeAttribute(attr));
  const groups = new Map();
  for (const r of document.querySelectorAll(rowSel)) {
    const t = r.closest('table');
    if (!t) continue;
    if (!groups.has(t)) groups.set(t, []);
    groups.get(t).push(r);
  }
  const tables = Array.from(groups.keys());
  const text = el => (el.innerText || '').trim();
  const containers = [];
  const out = tables.map((t, i) => {
    // Climb to the grid's container, but never so far that it holds another grid.
    let c = t;
    for (let k = 0; k < 8 && c.parentElement && c.parentElement !== document.body; k++) {
      const p = c.parentElement;
      if (tables.some(o => o !== t && p.contains(o))) break;
      c = p;
      if (/gridbox/i.test(String(c.className || ''))) break;
    }
    const key = prefix + i;
    c.setAttribute(attr, key);
    containers.push(c);

    const rows = groups.get(t);
    const matrix = rows.map(r => Array.from(r.cells).map(text));
    const cols = Math.max(0, ...matrix.map(r => r.length));

    // Header: possibly several rows with merged cells (the portal has
    // "Mark" over First/Second/Third). Lay the cells out on a grid, honouring
    // rowspan/colspan, and give each column its lowest non-empty name.
    const headerLabels = ht => {
      const grid = [];
      Array.from(ht.rows).forEach((hr, r) => {
        grid[r] = grid[r] || [];
        let col = 0;
        for (const cell of hr.cells) {
          while (grid[r][col] !== undefined) col++;
          const label = text(cell).replace(/\s+/g, ' ');
          for (let dr = 0; dr < (cell.rowSpan || 1); dr++) {
            grid[r + dr] = grid[r + dr] || [];
            for (let dc = 0; dc < (cell.colSpan || 1); dc++) grid[r + dr][col + dc] = label;
          }
          col += cell.colSpan || 1;
        }
      });
      const width = Math.max(0, ...grid.map(row => row.length));
      const names = [];
      for (let k = 0; k < width; k++) {
        let name = '';
        for (const row of grid) if (row[k]) name = row[k];
        names.push(name);
      }
      return names;
    };
    let labels = null;
    const headerTables = Array.from(c.querySelectorAll('table'))
      .filter(ht => ht !== t && !t.contains(ht) && !ht.contains(t))
      .sort((x, y) => /\bhdr\b/.test(y.className) - /\bhdr\b/.test(x.className));
    for (const ht of headerTables) {
      const names = headerLabels(ht);
      // A header row may have one extra/missing hidden column; real names must
      // cover most of the table.
      if (Math.abs(names.length - cols) > 1) continue;
      const fitted = Array.from({length: cols}, (_, k) => names[k] || '');
      if (fitted.filter(Boolean).length >= Math.min(2, cols)) { labels = fitted; break; }
    }

    const rect = c.getBoundingClientRect();
    const style = getComputedStyle(c);
    const visible = rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden';
    let onTop = false;
    if (visible) {
      const x = Math.min(Math.max(rect.left + rect.width / 2, 0), innerWidth - 1);
      const y = Math.min(Math.max(rect.top + Math.min(rect.height / 2, 80), 0), innerHeight - 1);
      const hit = document.elementFromPoint(x, y);
      onTop = !!hit && c.contains(hit);
    }
    let z = 0;
    for (let el = c; el && el !== document.documentElement; el = el.parentElement) {
      const v = parseInt(getComputedStyle(el).zIndex, 10);
      if (!isNaN(v)) z = Math.max(z, v);
    }
    return {key, rows: rows.length, cols, labels, visible, on_top: onTop, z, order: i, matrix};
  });
  // Only now are all tables tagged, so a title can't be credited to the wrong one.
  const docTitle = (document.title || '').toLowerCase();
  out.forEach((o, i) => {
    o.titled = gfTitled(containers[i], attr, needle) || (out.length === 1 && docTitle.includes(needle));
  });
  return out;
}
""" + ")"


@dataclass
class GridInfo:
    """One grade-style table found in the browser."""
    page: Page
    frame: Frame
    key: str
    page_index: int
    in_frame: bool
    rows: int
    cols: int
    labels: Optional[list[str]]
    visible: bool
    on_top: bool
    z: int
    order: int
    titled: bool = False  # sits in the portal's "Mark Entry" popup
    matrix: list[list[str]] = field(repr=False, default_factory=list)

    def priority(self):
        # The "Mark Entry" popup first; then visible and in front, newer
        # windows, embedded frames, higher stacking, later in the page.
        return (self.titled, self.visible, self.on_top, self.page_index, self.in_frame, self.z, self.order)

    def root(self):
        return self.frame.locator(f'[{GRID_ATTR}="{self.key}"]')

    def row_locator(self):
        return self.root().locator(ROW_LOCATOR)

    def describe(self) -> str:
        if self.labels:
            cols = ", ".join(self.labels[:5]) + ("…" if len(self.labels) > 5 else "")
        else:
            first = next((r for r in self.matrix if any(r)), [])
            cols = "e.g. " + ", ".join(v for v in first if v)[:60]
        front = " — in front" if self.on_top else ""
        return f"{self.rows} rows: {cols}{front}"


def _frame_is_titled(frame: Frame) -> bool:
    """Whether the <iframe> holding this frame sits under the popup title."""
    try:
        element = frame.frame_element()
        return bool(element.evaluate(
            "(el, [attr, needle]) => {" + _TITLE_HELPER_JS + " return gfTitled(el, attr, needle); }",
            [GRID_ATTR, POPUP_TITLE],
        ))
    except PwError:
        return False


_READ_ROWS_JS = "rows => rows.map(r => Array.from(r.cells).map(c => (c.innerText || '').trim()))"


def read_grid_matrix(grid: GridInfo) -> list[list[str]]:
    """All rows of one grid as lists of cell text."""
    return grid.row_locator().evaluate_all(_READ_ROWS_JS)


def scan_grid(grid: GridInfo) -> GridScan:
    """Columns, labels and current values of one discovered grid.

    If header labels can't be read, columns get generic "Column N" labels;
    the UI shows example values from each column so they can still be told
    apart without typing anything.
    """
    matrix = grid.matrix

    def col_values(i):
        return [r[i] if i < len(r) else "" for r in matrix]

    labels = grid.labels
    columns = [
        GridColumn(
            index=i,
            label=(labels[i] if labels and i < len(labels) else "") or f"Column {i + 1}",
            values=col_values(i),
        )
        for i in range(grid.cols)
    ]
    return GridScan(columns=columns, labels_found=labels is not None)


# ---------- Excel ----------

def list_sheets(path: str) -> list[str]:
    with pd.ExcelFile(path) as xls:
        return list(xls.sheet_names)


def _decimals_in_format(number_format: str) -> Optional[tuple[int, int]]:
    """(required, optional) decimal places of an Excel number format like '0.00'.

    None means "General" (or a format we don't interpret): show the number
    the way Excel's General format does.
    """
    fmt = (number_format or "General").split(";")[0]
    fmt = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]|_.|\*.', "", fmt)  # quoted text, escapes, colors, padding
    if fmt.strip().lower() in ("general", "@", "") or "%" in fmt or re.search(r"[dmyhs]", fmt, re.I):
        return None
    m = re.search(r"[0#?,]*\.([0#?]+)|[0#?]", fmt)
    if not m:
        return None
    frac = m.group(1) or ""
    return frac.count("0"), sum(frac.count(ch) for ch in "#?")


def excel_display_text(value, number_format: str = "General"):
    """The text Excel shows for a cell, e.g. 13.8333 with format '0' -> '14'.

    Returns None for empty cells. Rounding is half away from zero on the
    decimal value, like Excel. Float noise such as 15.200000000000001 is
    removed.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return str(value).upper()
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return str(value)
        places = _decimals_in_format(number_format)
        if places is None:
            # General: at most ~10 significant digits, like a normal-width Excel column.
            text = f"{value:.10g}"
            if "e" in text:
                text = format(Decimal(text), "f")
        else:
            required, optional = places
            q = Decimal(repr(value)).quantize(Decimal(1).scaleb(-(required + optional)), rounding=ROUND_HALF_UP)
            text = format(q, "f")
            if optional and "." in text:
                whole, frac = text.split(".")
                frac = frac.rstrip("0").ljust(required, "0")
                text = f"{whole}.{frac}" if frac else whole
        return "0" if re.fullmatch(r"-0(\.0*)?", text) else text
    text = str(value).strip()
    return text or None


def load_sheet(path: str, sheet: str) -> pd.DataFrame:
    """Read a sheet as the text Excel displays, dropping fully empty rows/columns.

    Values are strings exactly as shown in Excel (number formats applied),
    so grades are copied as the teacher sees them. Column names follow
    pandas' conventions ("Unnamed: 4", duplicate "Quiz.1").
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        rows = [
            [excel_display_text(c.value, getattr(c, "number_format", "General")) for c in row]
            for row in ws.iter_rows()
        ]
    finally:
        wb.close()

    if not rows:
        return pd.DataFrame()
    width = max(len(r) for r in rows)
    rows = [r + [None] * (width - len(r)) for r in rows]

    names, seen = [], {}
    for i, h in enumerate(rows[0]):
        base = h if h else f"Unnamed: {i}"
        n = seen.get(base, 0)
        seen[base] = n + 1
        names.append(base if n == 0 else f"{base}.{n}")

    df = pd.DataFrame(rows[1:], columns=names, dtype=object)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return df.reset_index(drop=True)


# ---------- planning (what will be written) ----------

@dataclass
class CellWrite:
    mapping: ColumnMapping
    value: str  # typed exactly, e.g. "12.5"


@dataclass
class StudentPlan:
    sid: str
    excel_row: int  # 1-based row number as seen in Excel
    writes: list[CellWrite]


@dataclass
class FillPlan:
    students: list[StudentPlan] = field(default_factory=list)
    not_in_portal: list[str] = field(default_factory=list)
    missing_from_excel: list[str] = field(default_factory=list)
    duplicate_ids: list[str] = field(default_factory=list)
    invalid_cells: list[str] = field(default_factory=list)
    blank_cells: int = 0
    zero_cells: int = 0  # 0 in Excel, left as they are (zero_as_blank)
    rows_without_id: int = 0
    rounded: list[str] = field(default_factory=list)  # e.g. "Row 5 (2201…), 'Quiz': 12.3 → 12.5"

    @property
    def cell_count(self) -> int:
        return sum(len(s.writes) for s in self.students)


def build_fill_plan(job: FillJob, df: pd.DataFrame, portal_ids: Optional[list[str]]) -> FillPlan:
    """Decide exactly which cells will be written, before touching the portal."""
    plan = FillPlan()
    id_col = job.id_mapping.excel_column
    ids = df[id_col].map(normalize_id)

    counts = ids[ids != ""].value_counts()
    dupes = set(counts[counts > 1].index)
    plan.duplicate_ids = sorted(dupes)

    portal_set = {p for p in (normalize_id(x) for x in portal_ids or []) if p}
    excel_set = set(ids[ids != ""])
    if portal_ids is not None:
        plan.missing_from_excel = sorted(portal_set - excel_set)

    for pos, sid in enumerate(ids):
        excel_row = pos + 2  # +1 header row, +1 for 1-based
        if sid == "":
            plan.rows_without_id += 1
            continue
        if sid in dupes:
            continue  # ambiguous: never guess which row is right
        if portal_ids is not None and sid not in portal_set:
            plan.not_in_portal.append(sid)
            continue

        writes = []
        for m in job.grade_mappings:
            raw = df.at[pos, m.excel_column]
            try:
                value = parse_grade(raw)
            except ValueError:
                plan.invalid_cells.append(
                    f"Row {excel_row} ({sid}), '{m.excel_column}': {str(raw).strip()!r} is not a number"
                )
                continue
            if value is None:
                plan.blank_cells += 1
                continue
            if job.zero_as_blank and Decimal(value) == 0:
                plan.zero_cells += 1
                continue
            if job.round_up is not None:
                new = apply_round_up(value, job.round_up)
                if new != value:
                    plan.rounded.append(f"Row {excel_row} ({sid}), '{m.excel_column}': {value} → {new}")
                    value = new
            writes.append(CellWrite(m, value))

        if writes:
            plan.students.append(StudentPlan(sid, excel_row, writes))

    return plan


# ---------- filling ----------

def row_for_student(grid: GridInfo, sid: str, id_column_index: int):
    """The student's row, searched only inside the chosen grid."""
    td_n = id_column_index + 1  # nth-child is 1-based
    escaped = sid.replace("\\", "\\\\").replace('"', '\\"')
    return grid.row_locator().filter(
        has=grid.frame.locator(f'td:nth-child({td_n}):text-is("{escaped}")')
    ).first


def _cell_shows(text: str, value: str) -> bool:
    text = text.strip()
    if text == "":
        return False
    try:
        return float(text) == float(value)
    except ValueError:
        # Portal displays something non-numeric; can't verify more precisely.
        return True


def edit_cell_commit(page: Page, row, td_index: int, value: str, focus_index: int) -> None:
    """Double-click a DHTMLX cell, select-all, type, commit with Tab.

    Retries a few times since the grid is known to occasionally ignore the
    first edit attempt. Success means the cell now shows the new value (not
    just "something"), so a stale old value isn't mistaken for success.

    Before editing, the row's own (read-only) ID cell is clicked to close any
    editor left open by the previous Tab. This used to click the middle of
    the page, which with a popup open could land on the table behind it.
    """
    modifier = "Meta" if sys.platform == "darwin" else "Control"

    page.bring_to_front()
    row.locator("td").nth(focus_index).click()

    cell = row.locator("td").nth(td_index)
    cell.scroll_into_view_if_needed()

    for _ in range(3):
        cell.dblclick()
        page.wait_for_timeout(80)

        page.keyboard.press(f"{modifier}+A")
        page.keyboard.insert_text(str(value))
        page.keyboard.press("Tab")
        page.wait_for_timeout(150)

        if _cell_shows(cell.inner_text(), value):
            return

        page.wait_for_timeout(200)

    raise RuntimeError(f"cell still shows {cell.inner_text().strip()!r}")


@dataclass
class FillResult:
    updated: int = 0
    not_found: list[str] = field(default_factory=list)
    failed: list[str] = field(default_factory=list)
    stopped: bool = False


ProgressCallback = Callable[[int, int, str], None]
"""Called as (current_index, total, message) while filling."""


def run_fill(
    session: PortalSession,
    job: FillJob,
    plan: FillPlan,
    on_progress: ProgressCallback,
    should_stop: Callable[[], bool],
) -> FillResult:
    grid = session.grid
    try:
        still_there = grid is not None and not grid.page.is_closed() and grid.root().count() > 0
    except PwError:
        still_there = False
    if not still_there:
        raise RuntimeError(
            "The grade table was closed or reloaded in the browser. Open it again, go back to "
            "“Match Columns” and click “Read the browser again”, then come back and press Start."
        )
    page = grid.page
    result = FillResult()
    total = len(plan.students)

    for i, student in enumerate(plan.students, start=1):
        if should_stop():
            result.stopped = True
            on_progress(i - 1, total, "Stopped by user.")
            break

        row = row_for_student(grid, student.sid, job.id_mapping.portal_index)
        try:
            row.wait_for(timeout=6000)
        except PwTimeoutError:
            result.not_found.append(student.sid)
            on_progress(i, total, f"Not found in portal: {student.sid}")
            continue
        except PwError as e:
            raise RuntimeError(f"Browser error while looking for {student.sid}: {e}") from e

        failed_here = []
        for w in student.writes:
            try:
                edit_cell_commit(page, row, w.mapping.portal_index, w.value, job.id_mapping.portal_index)
            except (RuntimeError, PwTimeoutError) as e:
                failed_here.append(w.mapping.portal_label)
                result.failed.append(f"{student.sid} · {w.mapping.portal_label} = {w.value}: {e}")

        if failed_here:
            on_progress(i, total, f"Problem with {student.sid}: couldn't set {', '.join(failed_here)}")
        else:
            result.updated += 1
            summary = ", ".join(f"{w.mapping.portal_label}={w.value}" for w in student.writes)
            on_progress(i, total, f"Updated {student.sid}: {summary}")

        if job.pace_seconds:
            page.wait_for_timeout(int(job.pace_seconds * 1000))

    return result
