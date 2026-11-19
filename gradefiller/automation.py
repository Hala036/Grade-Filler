"""Playwright-driven automation against the Al-Quds grades portal.

All portal/browser interaction lives here so the UI layer never touches
Playwright directly. Nothing in this module is Tkinter- or Flet-specific.

Nothing here knows about specific courses or exams: the grid's columns are
discovered at runtime and the caller supplies any number of
Excel-column -> portal-column mappings.
"""
from __future__ import annotations

import math
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

import pandas as pd
from playwright.sync_api import (
    sync_playwright,
    Browser,
    BrowserContext,
    Page,
    TimeoutError as PwTimeoutError,
    Error as PwError,
)

PORTAL_LOGIN_URL = "https://portal.alquds.edu/en/login"
COURSE_PAGE_URL = "https://portal.alquds.edu/en/registration/marks/course_list"

# DHTMLX grid row selector. Rows come in two alternating classes.
ROW_LOCATOR = "tr.ev_dhx_web, tr.odd_dhx_web"

PACE_SECONDS_DEFAULT = 0.12


# ---------- value normalization ----------

def round_half_up(x: float) -> int:
    return int(math.floor(x + 0.5))


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


def parse_grade(x) -> Optional[int]:
    """Parse an Excel grade cell. None for blank; ValueError for non-numbers."""
    if _is_blank(x):
        return None
    f = float(str(x).strip())
    if not math.isfinite(f):
        raise ValueError(f"not a number: {x!r}")
    return round_half_up(f)


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
    blank_as_zero: bool = False
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

    def __init__(self, state_path: Path):
        self.state_path = state_path
        self._playwright = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

    def _alive(self) -> bool:
        try:
            return bool(self.browser and self.browser.is_connected() and self.context)
        except PwError:
            return False

    def _ensure_browser(self) -> None:
        if self._alive():
            if self.page is None or self.page.is_closed():
                pages = [p for p in self.context.pages if not p.is_closed()]
                self.page = pages[-1] if pages else self.context.new_page()
            return

        self.close()
        self._playwright = sync_playwright().start()
        self.browser = self._playwright.chromium.launch(headless=False)
        self.context = None
        if self.state_path.exists():
            try:
                self.context = self.browser.new_context(storage_state=str(self.state_path))
            except (PwError, ValueError, OSError):
                # Corrupt/old session file: start fresh, user just logs in again.
                self.context = None
        if self.context is None:
            self.context = self.browser.new_context()
        self.page = self.context.new_page()

    def open_login(self) -> None:
        self._ensure_browser()
        self.page.goto(PORTAL_LOGIN_URL, wait_until="domcontentloaded")
        self.page.bring_to_front()

    def open_course_list(self) -> None:
        self._ensure_browser()
        self.page.goto(COURSE_PAGE_URL, wait_until="domcontentloaded")
        self.page.bring_to_front()

    def save_session(self) -> None:
        if not self._alive():
            raise RuntimeError("The browser window was closed. Open it again and log in.")
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        self.context.storage_state(path=str(self.state_path))

    def find_grid_page(self, timeout_ms: int = 5000) -> bool:
        """Look for the grade grid in any open tab; make that tab current."""
        if not self._alive():
            return False
        deadline = time.monotonic() + timeout_ms / 1000
        while True:
            for p in reversed(self.context.pages):
                try:
                    if not p.is_closed() and p.locator(ROW_LOCATOR).count() > 0:
                        self.page = p
                        return True
                except PwError:
                    continue
            if time.monotonic() >= deadline:
                return False
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


# ---------- grid scanning ----------

# The row selector (tr.ev_dhx_web / tr.odd_dhx_web) is the classic dhtmlxGrid
# skin's row class naming. That grid renders its header in a separate
# scrollable container from the body, roughly:
#   div.gridbox_dhx_web
#     div.hdr > div/table > tr > td.hdrcell   (header, one row of cells)
#     div.obj > table > tr.ev_dhx_web/.odd_dhx_web > td   (body rows)
# Exact wrapper tag/nesting varies by dhtmlx build, so try several header
# selectors most-specific first.
_HEADER_SELECTORS = [
    "div.hdr td.hdrcell",
    "div.hdr td",
    "div.hdr th",
    "table.hdr td",
    "table.hdr th",
    "[class*='hdr'] td[class*='hdrcell']",
]

_READ_ROWS_JS = "rows => rows.map(r => Array.from(r.cells).map(c => (c.innerText || '').trim()))"


def read_grid_matrix(page: Page) -> list[list[str]]:
    """All grid rows as lists of cell text."""
    return page.locator(ROW_LOCATOR).evaluate_all(_READ_ROWS_JS)


def scan_grid(page: Page) -> GridScan:
    """Inspect the live DHTMLX grid: its columns, labels and current values.

    If header labels can't be read, columns get generic "Column N" labels;
    the UI shows example values from each column so they can still be told
    apart without typing anything.
    """
    matrix = read_grid_matrix(page)
    cell_count = max((len(r) for r in matrix), default=0)

    def col_values(i):
        return [r[i] if i < len(r) else "" for r in matrix]

    labels: Optional[list[str]] = None
    for sel in _HEADER_SELECTORS:
        loc = page.locator(sel)
        if loc.count() != cell_count:
            # Matched something that doesn't line up with the data columns.
            continue
        candidate = [t.strip() for t in loc.all_inner_texts()]
        if all(candidate):
            labels = candidate
            break

    columns = [
        GridColumn(
            index=i,
            label=labels[i] if labels else f"Column {i + 1}",
            values=col_values(i),
        )
        for i in range(cell_count)
    ]
    return GridScan(columns=columns, labels_found=labels is not None)


# ---------- Excel ----------

def list_sheets(path: str) -> list[str]:
    with pd.ExcelFile(path) as xls:
        return list(xls.sheet_names)


def load_sheet(path: str, sheet: str) -> pd.DataFrame:
    """Read a sheet with string column names, dropping fully empty rows/columns."""
    df = pd.read_excel(path, sheet_name=sheet)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df.reset_index(drop=True)


# ---------- planning (what will be written) ----------

@dataclass
class CellWrite:
    mapping: ColumnMapping
    value: int


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
    rows_without_id: int = 0

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
                if job.blank_as_zero:
                    value = 0
                else:
                    plan.blank_cells += 1
                    continue
            writes.append(CellWrite(m, value))

        if writes:
            plan.students.append(StudentPlan(sid, excel_row, writes))

    return plan


# ---------- filling ----------

def row_for_student(page: Page, sid: str, id_column_index: int):
    rows = page.locator(ROW_LOCATOR)
    td_n = id_column_index + 1  # nth-child is 1-based
    escaped = sid.replace("\\", "\\\\").replace('"', '\\"')
    return rows.filter(has=page.locator(f'td:nth-child({td_n}):text-is("{escaped}")')).first


def _cell_shows(text: str, value: int) -> bool:
    text = text.strip()
    if text == "":
        return False
    try:
        return float(text) == float(value)
    except ValueError:
        # Portal displays something non-numeric; can't verify more precisely.
        return True


def edit_cell_commit(page: Page, row, td_index: int, value: int) -> None:
    """Double-click a DHTMLX cell, select-all, type, commit with Tab.

    Retries a few times since the grid is known to occasionally ignore the
    first edit attempt. Success means the cell now shows the new value (not
    just "something"), so a stale old value isn't mistaken for success.
    """
    modifier = "Meta" if sys.platform == "darwin" else "Control"

    page.bring_to_front()
    page.locator("body").click()

    cell = row.locator("td").nth(td_index)
    cell.scroll_into_view_if_needed()

    for _ in range(3):
        cell.dblclick()
        page.wait_for_timeout(80)

        page.keyboard.press(f"{modifier}+A")
        page.keyboard.insert_text(str(int(value)))
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
    if not session.find_grid_page(timeout_ms=5000):
        raise RuntimeError(
            "The grade table isn't open in the browser anymore. "
            "Open it again, then press Start."
        )
    page = session.page
    result = FillResult()
    total = len(plan.students)

    for i, student in enumerate(plan.students, start=1):
        if should_stop():
            result.stopped = True
            on_progress(i - 1, total, "Stopped by user.")
            break

        row = row_for_student(page, student.sid, job.id_mapping.portal_index)
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
                edit_cell_commit(page, row, w.mapping.portal_index, w.value)
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
