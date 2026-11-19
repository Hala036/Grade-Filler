"""Playwright-driven automation against the Al-Quds grades portal.

All portal/browser interaction lives here so the UI layer never touches
Playwright directly. Nothing in this module is Tkinter- or Flet-specific.

Nothing here knows about specific courses or exams: the grid's columns are
discovered at runtime and the caller supplies any number of
Excel-column -> portal-column mappings.
"""
from __future__ import annotations

import math
import re
import sys
import time
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
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

PORTAL_LOGIN_URL = "https://portal.alquds.edu/en/login"
COURSE_PAGE_URL = "https://portal.alquds.edu/en/registration/marks/course_list"

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
        self.grid: Optional[GridInfo] = None  # the table chosen for filling

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

    def find_grids(self, timeout_ms: int = 5000) -> list["GridInfo"]:
        """Every grade-style table in every tab and frame, best guess first.

        The portal shows a list of groups as one table and opens the real
        grade table in a popup on top of it, so several tables are normal.
        The one in front (popup / newest window) is ranked first.
        """
        if not self._alive():
            return []
        deadline = time.monotonic() + timeout_ms / 1000
        while True:
            grids: list[GridInfo] = []
            for pi, p in enumerate(self.context.pages):
                if p.is_closed():
                    continue
                for fi, f in enumerate(p.frames):
                    try:
                        found = f.evaluate(_DISCOVER_JS, [ROW_LOCATOR, GRID_ATTR, f"p{pi}f{fi}g"])
                    except PwError:
                        continue  # frame navigated/detached mid-scan
                    for d in found:
                        grids.append(
                            GridInfo(page=p, frame=f, page_index=pi, in_frame=f != p.main_frame, **d)
                        )
            if grids or time.monotonic() >= deadline:
                grids.sort(key=GridInfo.priority, reverse=True)
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

_DISCOVER_JS = r"""
([rowSel, attr, prefix]) => {
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
  return tables.map((t, i) => {
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

    const rows = groups.get(t);
    const matrix = rows.map(r => Array.from(r.cells).map(text));
    const cols = Math.max(0, ...matrix.map(r => r.length));

    let labels = null;
    for (const ht of c.querySelectorAll('table')) {
      if (ht === t || t.contains(ht)) continue;
      for (const hr of Array.from(ht.rows).reverse()) {
        const texts = Array.from(hr.cells).map(text);
        if (texts.length === cols && texts.every(Boolean)) { labels = texts; break; }
      }
      if (labels) break;
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
}
"""


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
    matrix: list[list[str]] = field(repr=False, default_factory=list)

    def priority(self):
        # Visible and in front first; then newer windows, embedded frames,
        # higher stacking (popups), and finally later in the page.
        return (self.visible, self.on_top, self.page_index, self.in_frame, self.z, self.order)

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
            label=labels[i] if labels else f"Column {i + 1}",
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
                    value = "0"
                else:
                    plan.blank_cells += 1
                    continue
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
