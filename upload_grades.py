import math
import sys
import time
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PwTimeoutError


# =========================
# CONFIG
# =========================
STATE_PATH = "state.json"
GRADES_XLSX = "2251-1214311-sec1-Grades.xlsx"
SHEET_NAME = "2251-1214311-sec2"

COURSE_PAGE_URL = "https://portal.alquds.edu/en/registration/marks/course_list"

# Timing
PACE_SECONDS = 0.12

# =========================
# EXCEL COLUMN NAMES
# =========================
COL_ID = "ID number"
COL_FIRST = "First/25"
COL_SECOND = "Second/25"
COL_QUIZ = "Quiz/10"
COL_FINAL = "Final/40"
COL_TOTAL = "Total 100"

# =========================
# DHTMLX GRID SELECTORS / INDEXES
# =========================
ROW_LOCATOR = "tr.ev_dhx_web, tr.odd_dhx_web"

IDX_STUDENT_NUMBER = 1
IDX_FIRST = 3
IDX_SECOND = 4
IDX_QUIZ = 5
IDX_FINAL = 6
IDX_TOTAL = 7


# =========================
# UTILITIES
# =========================
def round_half_up(x: float) -> int:
    return int(math.floor(x + 0.5))


def edit_cell_commit(page, row, td_index: int, value: int):
    modifier = "Meta" if sys.platform == "darwin" else "Control"

    page.bring_to_front()
    page.locator("body").click()

    cell = row.locator("td").nth(td_index)
    cell.scroll_into_view_if_needed()

    for attempt in range(2):
        cell.dblclick()
        page.wait_for_timeout(80)

        # Select all + replace
        page.keyboard.press(f"{modifier}+A")

        # Use insertText (more reliable on macOS than type with IME/keyboard layouts)
        page.keyboard.insert_text(str(int(value)))

        # Commit like human
        page.keyboard.press("Tab")
        page.wait_for_timeout(150)

        shown = cell.inner_text().strip()

        # If it updated, we're done
        if shown != "":
            return

        # Otherwise retry after a small pause
        page.wait_for_timeout(200)

    raise RuntimeError(f"Edit failed: cell stayed blank after 2 attempts (td={td_index}, value={value}).")


def row_for_student(page, sid: str):
    rows = page.locator(ROW_LOCATOR)

    # Exact match: td:nth-child(2) equals sid
    # (IDX_STUDENT_NUMBER is 1, +1 for nth-child, +1 because nth-child is 1-based => 2)
    try:
        return rows.filter(
            has=page.locator(f'td:nth-child({IDX_STUDENT_NUMBER + 1}):text-is("{sid}")')
        ).first
    except Exception:
        return rows.filter(
            has=page.locator(f"td:nth-child({IDX_STUDENT_NUMBER + 1})", has_text=sid)
        ).first


def require_excel_columns(df: pd.DataFrame, cols: list[str]):
    missing_cols = [c for c in cols if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Excel is missing columns: {missing_cols}")


# =========================
# MAIN
# =========================
def main():
    if not Path(STATE_PATH).exists():
        raise RuntimeError(f"{STATE_PATH} not found. Run your login/session saver first.")

    df = pd.read_excel(GRADES_XLSX, sheet_name=SHEET_NAME)

    required = [COL_ID, COL_FIRST, COL_SECOND, COL_QUIZ, COL_FINAL, COL_TOTAL]
    require_excel_columns(df, required)

    # Keep only needed columns
    df = df[required].copy()

    # Normalize ID: 21911103.0 -> "21911103"
    df[COL_ID] = df[COL_ID].apply(
        lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() != "" else ""
    ).str.strip()

    # Normalize grades to ints, blanks -> 0, round .5 up
    for c in [COL_FIRST, COL_SECOND, COL_QUIZ, COL_FINAL, COL_TOTAL]:
        df[c] = df[c].apply(
            lambda v: 0 if pd.isna(v) or str(v).strip() == "" else round_half_up(float(v))
        )

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state=STATE_PATH)
        page = context.new_page()

        page.goto(COURSE_PAGE_URL, wait_until="domcontentloaded")
        input("Navigate to the correct course/section grades table, then press Enter...")

        print("Waiting for grade grid rows to appear...")
        try:
            page.locator(ROW_LOCATOR).first.wait_for(timeout=60_000)
        except PwTimeoutError:
            raise RuntimeError("Grid rows never appeared. Make sure you're on the grades table page.")

        print("Detected grid rows:", page.locator(ROW_LOCATOR).count())

        for _, r in df.iterrows():
            sid = r[COL_ID]

            row = row_for_student(page, sid)
            try:
                row.wait_for(timeout=6000)
            except PwTimeoutError:
                print(f"[MISSING] {sid} not found in portal list")
                continue

            print(f"[UPDATING] {sid}: {r[COL_FIRST]}, {r[COL_SECOND]}, {r[COL_QUIZ]}, {r[COL_FINAL]}")
            edit_cell_commit(page, row, IDX_FIRST,  int(r[COL_FIRST]));  time.sleep(0.05)
            edit_cell_commit(page, row, IDX_SECOND, int(r[COL_SECOND])); time.sleep(0.05)
            edit_cell_commit(page, row, IDX_QUIZ,   int(r[COL_QUIZ]));   time.sleep(0.05)
            edit_cell_commit(page, row, IDX_FINAL,  int(r[COL_FINAL]))

            time.sleep(PACE_SECONDS)
        browser.close()


if __name__ == "__main__":
    main()
