"""Grade Filler wizard UI (Flet).

A single guided flow: log in -> pick Excel file -> open the course grid ->
map columns (scanned live from the portal) -> review -> fill -> save.

Works for any course and any number of exams: the portal's columns are
read from the live grade table, the Excel columns from the chosen sheet,
and the user maps as many Excel -> portal columns as they need (the app
pre-fills the obvious ones).

Playwright's sync API pins itself to whichever OS thread first calls
sync_playwright().start() — every later call for that session must run on
the *same* thread or Playwright raises "cannot switch to a different
thread". Flet button clicks and dialog confirmations aren't guaranteed to
land on the same thread as each other, so all Playwright calls are routed
through a single PlaywrightWorker (gradefiller/playwright_worker.py) that
owns one dedicated thread for the whole browser session's lifetime.
"""
from __future__ import annotations

import logging
import threading
from logging.handlers import RotatingFileHandler
import time
import traceback
from typing import Optional

import flet as ft
import pandas as pd

from gradefiller import automation, matching
from gradefiller.round_line import RoundLine, tenth_label
from gradefiller.paths import app_data_dir, state_path
from gradefiller.playwright_worker import PlaywrightWorker

log = logging.getLogger("gradefiller")


def _ms(since: float) -> float:
    return (time.perf_counter() - since) * 1000


def setup_logging() -> None:
    """Timing/diagnostic lines go only to GradeFiller/gradefiller.log, not the terminal.

    The file is capped (about 1 MB, one old copy kept). To see the lines live
    while debugging, add a logging.StreamHandler() to `log` here.
    """
    if log.handlers:
        return
    log.setLevel(logging.INFO)
    log.propagate = False  # never fall through to the terminal
    try:
        file = RotatingFileHandler(
            app_data_dir() / "gradefiller.log", maxBytes=1_000_000, backupCount=1, encoding="utf-8"
        )
        file.setFormatter(logging.Formatter("%(asctime)s %(message)s", "%Y-%m-%d %H:%M:%S"))
        log.addHandler(file)
    except OSError:
        log.addHandler(logging.NullHandler())  # logging must never stop the app

STEP_LOGIN, STEP_EXCEL, STEP_GRID, STEP_MAP, STEP_RUN = range(5)

ERROR = ft.Colors.RED_700
OK = ft.Colors.GREEN_700
INFO = ft.Colors.BLUE_GREY_700
WARN = ft.Colors.ORANGE_800


def _preview(values: list[str], limit: int = 8) -> str:
    shown = ", ".join(values[:limit])
    extra = len(values) - limit
    return f"{shown} and {extra} more" if extra > 0 else shown


def _fmt(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:.2f}".rstrip("0").rstrip(".")
    return str(v).strip()


class GradeFillerApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Grade Filler"
        self.page.window.width = 900
        self.page.window.height = 720
        self.page.window.min_width = 720
        self.page.window.min_height = 600
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.padding = 0

        self.pw = PlaywrightWorker(dispatch=self._on_ui_thread)
        self.session: Optional[automation.PortalSession] = None
        self.file_picker = ft.FilePicker()  # keep a reference: services are dropped otherwise

        # Step 2
        self.excel_path: Optional[str] = None
        self.sheets: list[str] = []
        self.sheet_name: Optional[str] = None
        self.df: Optional[pd.DataFrame] = None

        # Step 3/4
        self.grid: list[automation.GridColumn] = []
        self.grid_choices: list[automation.GridInfo] = []  # all tables found, best first
        self.table_match: Optional[matching.GridMatch] = None
        self._show_table_picker = False  # expanded when unsure, or on "Change"
        self._show_id_editor = False
        self.labels_found = True
        self.id_portal: Optional[int] = None
        self.id_excel: Optional[str] = None
        self.grade_pairs: list[dict] = []  # {"excel": str|None, "portal": int|None}
        self.blank_as_zero = False
        self.round_up_on = False
        self.round_rule = automation.RoundUp()  # the usual rule: .1-.4 -> .5, .5+ -> 1
        self._mapping_source = None  # what the current mapping was suggested for

        # Step 5
        self.job: Optional[automation.FillJob] = None
        self.plan: Optional[automation.FillPlan] = None
        self._stop_requested = False
        self._running = False

        self.steps = [
            self._build_step_login,
            self._build_step_excel,
            self._build_step_grid,
            self._build_step_mapping,
            self._build_step_run,
        ]
        self.step_index = STEP_LOGIN

        self.page.window.prevent_close = True
        self.page.window.on_event = self._on_window_event

        self.body = ft.Container(expand=True, padding=24)
        self.header = self._build_header()
        self.page.add(
            ft.Column(
                [self.header, ft.Divider(height=1), self.body],
                expand=True,
                spacing=0,
            )
        )
        self._render_step()

    # ---------- shell ----------

    def _on_ui_thread(self, callback, *args):
        """Run callback(*args) on Flet's event loop, from any thread.

        Screen changes made from the browser thread were queued but not
        shown until something else woke Flet up (like switching windows).
        """
        async def run():
            callback(*args)

        try:
            self.page.run_task(run)
        except Exception:
            # No running event loop (e.g. during shutdown): best effort.
            callback(*args)

    def _session(self) -> automation.PortalSession:
        if self.session is None:
            self.session = automation.PortalSession(state_path())
        return self.session

    def _build_header(self) -> ft.Control:
        labels = ["Log In", "Grades File", "Course Table", "Match Columns", "Review & Fill"]
        self.step_chips = []
        for i, label in enumerate(labels):
            chip = ft.Container(
                content=ft.Text(f"{i + 1}. {label}", size=12, weight=ft.FontWeight.W_500),
                padding=ft.Padding.symmetric(horizontal=12, vertical=6),
                border_radius=16,
                on_click=lambda _, i=i: self._chip_clicked(i),
            )
            self.step_chips.append(chip)
        return ft.Container(
            content=ft.Row(self.step_chips, spacing=8, wrap=True),
            padding=ft.Padding.symmetric(horizontal=24, vertical=16),
        )

    def _chip_clicked(self, index: int):
        # Only allow jumping back to steps already done.
        if index < self.step_index and not self._running:
            self._go_to(index)

    def _update_chips(self):
        for i, chip in enumerate(self.step_chips):
            if i == self.step_index:
                chip.bgcolor = ft.Colors.BLUE_600
                chip.content.color = ft.Colors.WHITE
                chip.tooltip = None
            elif i < self.step_index:
                chip.bgcolor = ft.Colors.GREEN_100
                chip.content.color = None
                chip.tooltip = "Go back to this step"
            else:
                chip.bgcolor = ft.Colors.BLUE_GREY_50
                chip.content.color = None
                chip.tooltip = None

    def _render_step(self):
        self._update_chips()
        self.body.content = self.steps[self.step_index]()
        self.page.update()

    def _go_to(self, index: int):
        self.step_index = index
        self._render_step()

    def _step_layout(self, title: str, intro: str, controls: list, back_to: Optional[int] = None):
        top = [ft.Text(title, size=20, weight=ft.FontWeight.BOLD)]
        if intro:
            top.append(ft.Text(intro, size=14))
        top.append(ft.Container(height=8))
        bottom = []
        if back_to is not None:
            bottom = [
                ft.Container(height=8),
                ft.TextButton("Back", icon=ft.Icons.ARROW_BACK, on_click=lambda _: self._go_to(back_to)),
            ]
        return ft.Column(top + controls + bottom, spacing=10, scroll=ft.ScrollMode.AUTO)

    @staticmethod
    def _set_status(text: ft.Text, message: str, color=INFO):
        text.value = message
        text.color = color

    def _error_text(self, e: Exception) -> str:
        msg = str(e).strip() or e.__class__.__name__
        first = msg.splitlines()[0]
        if "Target page, context or browser has been closed" in msg or "Browser has been closed" in msg:
            return "The browser window was closed. Click the button again to reopen it."
        return first

    # ---------- Step 1: Login ----------

    def _build_step_login(self) -> ft.Control:
        status = ft.Text("")
        has_saved = state_path().exists()

        confirm_btn = ft.FilledButton("I've logged in — continue", icon=ft.Icons.CHECK, visible=False)
        open_btn_cls = ft.OutlinedButton if has_saved else ft.FilledButton
        open_btn = open_btn_cls(
            "Log in again" if has_saved else "Open browser to log in",
            icon=ft.Icons.LOGIN,
        )

        def open_click(_):
            open_btn.disabled = True
            self._set_status(status, "Opening the browser…")
            self.page.update()

            def ok(_r):
                open_btn.disabled = False
                confirm_btn.visible = True
                confirm_btn.disabled = False
                self._set_status(
                    status,
                    "Log in in the browser window, then come back and click “I've logged in — continue”.",
                )
                self.page.update()

            def err(e):
                open_btn.disabled = False
                self._set_status(status, f"Couldn't open the browser: {self._error_text(e)}", ERROR)
                self.page.update()

            self.pw.submit(self._session().open_login, ok, err)

        def confirm_click(_):
            confirm_btn.disabled = True
            self._set_status(status, "Saving your login…")
            self.page.update()

            def err(e):
                confirm_btn.disabled = False
                self._set_status(status, self._error_text(e), ERROR)
                self.page.update()

            self.pw.submit(self._session().save_session, lambda _r: self._go_to(STEP_EXCEL), err)

        open_btn.on_click = open_click
        confirm_btn.on_click = confirm_click

        controls = []
        if has_saved:
            controls += [
                ft.Row(
                    [
                        ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_600),
                        ft.Text("You logged in on this computer before, so you can probably skip this."),
                    ]
                ),
                ft.Row(
                    [
                        ft.FilledButton(
                            "Continue with saved login",
                            icon=ft.Icons.ARROW_FORWARD,
                            on_click=lambda _: self._go_to(STEP_EXCEL),
                        ),
                        open_btn,
                    ],
                    wrap=True,
                ),
                ft.Text(
                    "If the saved login has expired, the portal will just ask you to log in again later.",
                    size=12,
                    color=INFO,
                ),
            ]
        else:
            controls.append(open_btn)
        controls += [confirm_btn, status]

        return self._step_layout(
            "Step 1: Log in to the portal",
            "This opens a browser window. Log in as usual — your login is remembered "
            "on this computer so you won't need to do it every time.",
            controls,
        )

    # ---------- Step 2: Excel file ----------

    def _build_step_excel(self) -> ft.Control:
        path_text = ft.Text(self.excel_path or "No file selected", italic=self.excel_path is None, expand=True)
        status = ft.Text("")
        sheet_dropdown = ft.Dropdown(
            label="Sheet",
            width=320,
            options=[ft.dropdown.Option(s) for s in self.sheets],
            value=self.sheet_name,
            visible=bool(self.sheets),
        )
        preview_holder = ft.Column([])
        next_btn = ft.FilledButton(
            "Continue",
            icon=ft.Icons.ARROW_FORWARD,
            disabled=self.df is None or self.df.empty,
            on_click=lambda _: self._go_to(STEP_GRID),
        )

        def show_preview():
            df = self.df
            if df is None:
                preview_holder.controls = []
                return
            if df.empty:
                preview_holder.controls = [ft.Text("This sheet has no data rows.", color=ERROR)]
                return
            head = df.head(5)
            table = ft.DataTable(
                columns=[ft.DataColumn(ft.Text(c, weight=ft.FontWeight.BOLD)) for c in df.columns],
                rows=[
                    ft.DataRow(cells=[ft.DataCell(ft.Text(_fmt(v))) for v in row])
                    for row in head.itertuples(index=False)
                ],
                heading_row_height=36,
                data_row_min_height=30,
                data_row_max_height=30,
                column_spacing=18,
            )
            preview_holder.controls = [
                ft.Text(
                    f"{len(df)} rows, {len(df.columns)} columns. First 5 rows:",
                    weight=ft.FontWeight.W_500,
                ),
                ft.Container(
                    content=ft.Row([table], scroll=ft.ScrollMode.AUTO),
                    border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
                    border_radius=8,
                ),
            ]

        def load_sheet():
            try:
                self.df = automation.load_sheet(self.excel_path, self.sheet_name)
                self._set_status(status, "")
            except Exception as e:
                self.df = None
                self._set_status(status, f"Couldn't read this sheet: {self._error_text(e)}", ERROR)
            next_btn.disabled = self.df is None or self.df.empty
            show_preview()
            self.page.update()

        def on_sheet_select(e):
            self.sheet_name = sheet_dropdown.value
            load_sheet()

        async def pick_file(_):
            files = await self.file_picker.pick_files(
                dialog_title="Choose the Excel grades file",
                allowed_extensions=["xlsx", "xlsm"],
                allow_multiple=False,
            )
            if not files:
                return
            path = files[0].path
            if not path:
                self._set_status(status, "Couldn't get the file's location. Try choosing it again.", ERROR)
                self.page.update()
                return
            try:
                sheets = automation.list_sheets(path)
            except Exception as e:
                self._set_status(
                    status,
                    f"Couldn't open this file: {self._error_text(e)}. If it's open in Excel, "
                    "close it and try again.",
                    ERROR,
                )
                self.page.update()
                return
            if not sheets:
                self._set_status(status, "This file has no sheets.", ERROR)
                self.page.update()
                return

            self.excel_path = path
            self.sheets = sheets
            self.sheet_name = sheets[0]
            path_text.value = path
            path_text.italic = False
            sheet_dropdown.options = [ft.dropdown.Option(s) for s in sheets]
            sheet_dropdown.value = self.sheet_name
            sheet_dropdown.visible = True
            load_sheet()

        sheet_dropdown.on_select = on_sheet_select
        show_preview()

        return self._step_layout(
            "Step 2: Choose your grades file",
            "Pick the Excel file with the grades, then the sheet for this section. "
            "The first row of the sheet should be the column names.",
            [
                ft.Row(
                    [
                        ft.FilledButton(
                            "Choose Excel file…" if not self.excel_path else "Choose a different file…",
                            icon=ft.Icons.UPLOAD_FILE,
                            on_click=pick_file,
                        ),
                        path_text,
                    ]
                ),
                sheet_dropdown,
                status,
                preview_holder,
                ft.Container(height=8),
                next_btn,
            ],
            back_to=STEP_LOGIN,
        )

    # ---------- Step 3: Open the course grade table ----------

    def _build_step_grid(self) -> ft.Control:
        status = ft.Text("")
        open_btn = ft.OutlinedButton("Open the course list in the browser", icon=ft.Icons.OPEN_IN_BROWSER)
        ready_btn = ft.FilledButton("The grade table is showing — continue", icon=ft.Icons.ARROW_FORWARD)

        def open_click(_):
            open_btn.disabled = True
            self._set_status(status, "Opening…")
            self.page.update()

            def ok(_r):
                open_btn.disabled = False
                self._set_status(
                    status,
                    "Now click through to the course, section and grade table in the browser. "
                    "If the portal asks you to log in, do that first.",
                )
                self.page.update()

            def err(e):
                open_btn.disabled = False
                self._set_status(status, f"Couldn't open the portal: {self._error_text(e)}", ERROR)
                self.page.update()

            self.pw.submit(self._session().open_course_list, ok, err)

        def ready_click(_):
            ready_btn.disabled = True
            open_btn.disabled = True
            self._set_status(status, "Looking for the grade table…")
            self.page.update()
            self._scan(
                on_done=lambda: self._go_to(STEP_MAP),
                on_fail=lambda msg: (
                    setattr(ready_btn, "disabled", False),
                    setattr(open_btn, "disabled", False),
                    self._set_status(status, msg, ERROR),
                    self.page.update(),
                ),
            )

        open_btn.on_click = open_click
        ready_btn.on_click = ready_click

        return self._step_layout(
            "Step 3: Open the grade table",
            "In the browser, go to the exact course and section you want to fill and open "
            "its grade table (if it opens in a popup, leave the popup open). Every course's "
            "table is different, so this part is up to you. Then come back here.",
            [ft.Row([open_btn, ready_btn], wrap=True), status],
            back_to=STEP_EXCEL,
        )

    def _scan(self, on_done, on_fail):
        """Find the grade table in the browser and read its columns."""

        clicked = time.perf_counter()

        def task():
            log.info("Looking for the grade table (waited %.0f ms for the browser to be free)", _ms(clicked))
            session = self._session()
            grids = session.find_grids(timeout_ms=5000)
            t = time.perf_counter()
            try:
                session.save_session()  # keep the saved login fresh
            except Exception:
                pass
            log.info("save_session: %.0f ms", _ms(t))
            return grids

        def ok(grids):
            grids = [g for g in grids if g.cols > 0]
            if not grids:
                log.info("No tables found (%.0f ms total)", _ms(clicked))
                on_fail(
                    "Couldn't find a student grade table in the browser. Make sure it's open "
                    "and shows the students, then try again."
                )
                return
            t = time.perf_counter()
            ranked, confident = matching.rank_grids(grids, self.df)
            titled = [m for m in ranked if m.grid.titled]
            if len(titled) == 1:
                # The "Mark Entry" popup is the table the user opened on purpose.
                ranked = titled + [m for m in ranked if m is not titled[0]]
                confident = True
                how = "by the 'Mark Entry' popup title"
            elif titled:
                how = f"by student IDs ({len(titled)} tables looked like Mark Entry)"
            else:
                how = "by student IDs (no 'Mark Entry' popup found)"
            log.info("Table chosen %s: %d students; %s; ranking %.0f ms",
                     how, ranked[0].grid.rows, "no question needed" if confident else "user must pick", _ms(t))
            self.grid_choices = [m.grid for m in ranked]
            self.table_match = ranked[0]
            # Ask only when there's a real choice and the data didn't settle it.
            self._show_table_picker = len(ranked) > 1 and not confident
            self._show_id_editor = False
            self._use_grid(ranked[0].grid)
            t = time.perf_counter()
            on_done()
            log.info("Next screen built in %.0f ms; %.0f ms in total since the click", _ms(t), _ms(clicked))

        def err(e):
            log.info("Table search failed after %.0f ms: %s", _ms(clicked), e)
            on_fail(f"Couldn't read the grade table: {self._error_text(e)}")

        self.pw.submit(task, ok, err)

    def _use_grid(self, info: automation.GridInfo):
        """Make this browser table the one to read and fill."""
        self._session().grid = info
        scan = automation.scan_grid(info)
        self.grid = scan.columns
        self.labels_found = scan.labels_found

    # ---------- Step 4: Column mapping ----------

    def _grid_col(self, index: Optional[int]) -> Optional[automation.GridColumn]:
        return next((g for g in self.grid if g.index == index), None)

    def _suggest_mapping(self):
        """Pre-fill ID and grade pairs. Only when file/sheet/table changed."""
        chosen = self._session().grid
        source = (
            self.excel_path,
            self.sheet_name,
            chosen.key if chosen else None,
            tuple(g.label for g in self.grid),
        )
        if source == self._mapping_source:
            # Same data: keep the user's choices, just drop ones that no longer exist.
            self._prune_mapping()
            return
        self._mapping_source = source

        p, e, n = matching.detect_id_columns(self.grid, self.df)
        self.id_portal, self.id_excel = (p, e) if n > 0 else (None, None)

        portal_cols = matching.portal_fillable_columns(self.grid, self.id_portal)
        excel_cols = matching.grade_candidate_columns(self.df, self.id_excel)
        pairs = matching.suggest_grade_mappings(portal_cols, excel_cols)
        self.grade_pairs = [{"portal": pi, "excel": ec} for pi, ec in pairs] or [{"portal": None, "excel": None}]

    def _prune_mapping(self):
        indexes = {g.index for g in self.grid}
        cols = set(self.df.columns)
        if self.id_portal not in indexes:
            self.id_portal = None
        if self.id_excel not in cols:
            self.id_excel = None
        for pair in self.grade_pairs:
            if pair["portal"] not in indexes:
                pair["portal"] = None
            if pair["excel"] not in cols:
                pair["excel"] = None

    def _mapping_problems(self) -> tuple[list[str], dict[int, str]]:
        """(general problems, per-row problems by row number)."""
        general, per_row = [], {}
        if self.id_portal is None or self.id_excel is None:
            general.append("Choose the Student ID column on both sides.")

        seen_portal: dict[int, int] = {}
        any_complete = False
        for n, pair in enumerate(self.grade_pairs):
            p, e = pair["portal"], pair["excel"]
            if p is None and e is None:
                continue
            if p is None or e is None:
                per_row[n] = "Choose both columns, or remove this row."
                continue
            if p == self.id_portal:
                per_row[n] = "That's the Student ID column — it can't be filled."
                continue
            if p in seen_portal:
                per_row[n] = f"This portal column is already used in row {seen_portal[p] + 1}."
                continue
            seen_portal[p] = n
            any_complete = True
        if not any_complete and not per_row:
            general.append("Add at least one grade column to fill.")
        return general, per_row

    def _round_up_rule(self) -> tuple[Optional[automation.RoundUp], Optional[str]]:
        """(rule, problem). Rule is None when rounding is off or nothing is chosen."""
        if not self.round_up_on:
            return None, None
        if not self.round_rule.active:
            return None, "Rounding is on, but no arrow is showing. Drag the dots apart, or untick it."
        return self.round_rule, None

    @staticmethod
    def _round_summary(rule: automation.RoundUp) -> str:
        """The rule in words, e.g. "below 0.2 stays · 0.2–0.49 → 0.5 · 0.5–0.99 → 1"."""
        def upto(n: int) -> str:  # the highest value under an arrow ending at n
            return f"0.{n - 1}9"

        parts = []
        if rule.start > 0:
            parts.append(f"below {tenth_label(rule.start)} stays")
        if rule.start < rule.joint:
            parts.append(f"{tenth_label(rule.start)}–{upto(rule.joint)} → {tenth_label(rule.joint)}")
        if rule.joint < rule.end:
            parts.append(f"{tenth_label(rule.joint)}–{upto(rule.end)} → {tenth_label(rule.end)}")
        if rule.end < 10:
            parts.append(f"{tenth_label(rule.end)} and above stays")
        return " · ".join(parts)

    def _build_step_mapping(self) -> ft.Control:
        self._suggest_mapping()
        holder = ft.Column([], spacing=10)
        status = ft.Text("")

        def portal_option(g: automation.GridColumn) -> ft.dropdown.Option:
            sample = g.sample_value
            hint = f"  (now: {sample})" if sample else "  (empty)"
            return ft.dropdown.Option(key=str(g.index), text=f"{g.label}{hint}")

        def excel_option(col: str) -> ft.dropdown.Option:
            sample = next((_fmt(v) for v in self.df[col] if _fmt(v)), "")
            return ft.dropdown.Option(key=col, text=f"{col}  (e.g. {sample})" if sample else col)

        def render():
            general, per_row = self._mapping_problems()

            # --- Student ID card ---
            id_portal_dd = ft.Dropdown(
                label="Student ID column in the portal",
                options=[portal_option(g) for g in self.grid],
                value=str(self.id_portal) if self.id_portal is not None else None,
                expand=True,
            )
            id_excel_dd = ft.Dropdown(
                label="Student ID column in Excel",
                options=[excel_option(c) for c in self.df.columns],
                value=self.id_excel,
                expand=True,
            )

            def on_id_portal(e):
                self.id_portal = int(e.control.value) if e.control.value is not None else None
                render()

            def on_id_excel(e):
                self.id_excel = e.control.value
                render()

            id_portal_dd.on_select = on_id_portal
            id_excel_dd.on_select = on_id_excel

            found = total = 0
            if self.id_portal is not None and self.id_excel is not None:
                found, total = matching.id_overlap(self._grid_col(self.id_portal), self.df, self.id_excel)
            id_ok = found > 0
            # Few matches usually means the sheet is for another section/group.
            few = id_ok and found < 0.5 * total

            def toggle_id_editor(_):
                self._show_id_editor = not self._show_id_editor
                render()

            if id_ok and not self._show_id_editor:
                # Matched automatically: one line instead of two dropdowns.
                id_section = self._fact(
                    ft.Icons.WARNING_AMBER if few else ft.Icons.CHECK_CIRCLE,
                    f"Only {found} of {total} students are in your sheet. Is it the right sheet for this group?"
                    if few else f"{found} of {total} students found in your sheet.",
                    WARN if few else OK,
                    action=ft.TextButton("Change", on_click=toggle_id_editor),
                )
            else:
                if id_ok:
                    msg, color = f"✓ {found} of {total} students found in your sheet.", OK
                else:
                    msg, color = (
                        "Choose the student ID column on each side. (No matching student IDs were found. "
                        "Is this the right sheet and table?)",
                        ERROR,
                    )
                id_section = ft.Column(
                    [
                        ft.Text(msg, color=color, size=13),
                        ft.Row([id_portal_dd, id_excel_dd], spacing=12),
                    ]
                    + ([ft.TextButton("Done", on_click=toggle_id_editor)] if id_ok else []),
                    spacing=8,
                )

            # --- Grade rows card ---
            portal_choices = matching.portal_fillable_columns(self.grid, self.id_portal)
            # Keep currently chosen portal columns selectable even if filtered.
            chosen = {p["portal"] for p in self.grade_pairs if p["portal"] is not None}
            portal_choices += [g for g in self.grid if g.index in chosen and g not in portal_choices]
            portal_choices.sort(key=lambda g: g.index)

            excel_choices = matching.grade_candidate_columns(self.df, self.id_excel)
            chosen_e = [p["excel"] for p in self.grade_pairs if p["excel"] and p["excel"] not in excel_choices]
            excel_choices = [c for c in self.df.columns if c in set(excel_choices) | set(chosen_e)]

            rows = []
            for n, pair in enumerate(self.grade_pairs):
                excel_dd = ft.Dropdown(
                    label="Take grades from (Excel)",
                    options=[excel_option(c) for c in excel_choices],
                    value=pair["excel"],
                    expand=True,
                )
                portal_dd = ft.Dropdown(
                    label="Write them into (portal)",
                    options=[portal_option(g) for g in portal_choices],
                    value=str(pair["portal"]) if pair["portal"] is not None else None,
                    expand=True,
                )

                def on_excel(e, pair=pair):
                    pair["excel"] = e.control.value
                    render()

                def on_portal(e, pair=pair):
                    pair["portal"] = int(e.control.value) if e.control.value is not None else None
                    render()

                def remove(_, n=n):
                    del self.grade_pairs[n]
                    if not self.grade_pairs:
                        self.grade_pairs.append({"portal": None, "excel": None})
                    render()

                excel_dd.on_select = on_excel
                portal_dd.on_select = on_portal
                row_controls = [
                    ft.Row(
                        [
                            excel_dd,
                            ft.Icon(ft.Icons.ARROW_FORWARD, color=ft.Colors.BLUE_GREY_400),
                            portal_dd,
                            ft.IconButton(ft.Icons.DELETE_OUTLINE, tooltip="Remove this row", on_click=remove),
                        ],
                        spacing=8,
                    )
                ]
                if n in per_row:
                    row_controls.append(ft.Text(per_row[n], color=ERROR, size=12))
                rows.append(ft.Column(row_controls, spacing=2))

            def add_row(_):
                self.grade_pairs.append({"portal": None, "excel": None})
                render()

            blank_cb = ft.Checkbox(
                label="Write 0 for blank Excel cells (otherwise they're left as they are)",
                value=self.blank_as_zero,
            )

            def on_blank(e):
                self.blank_as_zero = bool(e.control.value)

            blank_cb.on_change = on_blank

            # --- Optional round-up: two joined arrows on a 0-1 line ---
            round_cb = ft.Checkbox(label="Round grades up", value=self.round_up_on)

            def on_round(e):
                self.round_up_on = bool(e.control.value)
                render()

            round_cb.on_change = on_round
            round_controls = [round_cb]
            if self.round_up_on:
                round_text = ft.Text("", size=12)

                def show_round_text():
                    _, problem = self._round_up_rule()
                    round_text.value = problem or (
                        self._round_summary(self.round_rule)
                        + ". This is the part after the point; whole numbers stay as they are."
                    )
                    round_text.color = ERROR if problem else INFO

                def on_line_change(rule):
                    self.round_rule = rule
                    show_round_text()
                    round_text.update()

                show_round_text()
                round_controls.append(
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Text(
                                    "Grades under an arrow are rounded up to where it points. "
                                    "Drag the dots, or click on the line.",
                                    size=13,
                                    color=INFO,
                                ),
                                RoundLine(self.round_rule, on_line_change).control,
                                round_text,
                            ],
                            spacing=6,
                        ),
                        padding=ft.Padding.only(left=32),
                    )
                )

            filled = sum(1 for p in self.grade_pairs if p["portal"] is not None and p["excel"])
            grades_card = self._card(
                "Grades to fill",
                [
                    ft.Text(
                        "Pre-filled where the names match. Please check them." if filled
                        else "Pick which Excel column goes into which portal column.",
                        size=13,
                        color=INFO,
                    ),
                    *rows,
                    ft.TextButton("Add another column", icon=ft.Icons.ADD, on_click=add_row),
                    blank_cb,
                    *round_controls,
                ],
            )

            notices = []
            current = self._session().grid

            def toggle_table_picker(_):
                self._show_table_picker = not self._show_table_picker
                render()

            if len(self.grid_choices) > 1 and self._show_table_picker:
                table_dd = ft.Dropdown(
                    label="Table to fill",
                    options=[
                        ft.dropdown.Option(key=g.key, text=f"Table {n + 1} · {g.describe()}")
                        for n, g in enumerate(self.grid_choices)
                    ],
                    value=current.key if current else None,
                    expand=True,
                )

                def on_table(e):
                    info = next((g for g in self.grid_choices if g.key == e.control.value), None)
                    if info:
                        self._use_grid(info)
                        self._suggest_mapping()
                        self._show_id_editor = False
                        render()

                table_dd.on_select = on_table
                notices.append(
                    ft.Column(
                        [
                            ft.Text(
                                "Several tables are open in the browser. Which one has these students' grades?",
                                color=WARN if not id_ok else INFO,
                                size=13,
                            ),
                            ft.Row([table_dd]),
                        ]
                        + ([ft.TextButton("Done", on_click=toggle_table_picker)] if id_ok else []),
                        spacing=8,
                    )
                )
            elif current is not None:
                notices.append(
                    self._fact(
                        ft.Icons.TABLE_CHART,
                        f"Filling the Mark Entry table ({current.rows} students)." if current.titled
                        else f"Filling the portal table with {current.rows} students.",
                        INFO,
                        action=ft.TextButton("Wrong table?", on_click=toggle_table_picker)
                        if len(self.grid_choices) > 1 else None,
                    )
                )
            notices.append(id_section)
            if not self.labels_found:
                notices.append(
                    ft.Text("Portal column names couldn't be read; they're numbered left to right.",
                            size=12, color=INFO)
                )

            def rescan(_):
                rescan_btn.disabled = True
                self._set_status(status, "Reading the grade table again…")
                self.page.update()

                def fail(msg):
                    rescan_btn.disabled = False
                    self._set_status(status, msg, ERROR)
                    self.page.update()

                self._scan(on_done=lambda: (self._set_status(status, ""), self._go_to(STEP_MAP)), on_fail=fail)

            rescan_btn = ft.TextButton(
                "Opened a different table? Read the browser again",
                icon=ft.Icons.REFRESH,
                on_click=rescan,
            )

            def go_next(_):
                general_now, per_row_now = self._mapping_problems()
                if general_now or per_row_now:
                    self._set_status(status, " ".join(general_now) or "Fix the rows marked in red above.", ERROR)
                    self.page.update()
                    return
                _, round_problem = self._round_up_rule()
                if round_problem:
                    self._set_status(status, round_problem, ERROR)
                    self.page.update()
                    return
                self._build_job_and_plan()
                self._go_to(STEP_RUN)

            holder.controls = notices + [
                grades_card,
                ft.Row([rescan_btn]),
                status,
                ft.FilledButton("Review", icon=ft.Icons.ARROW_FORWARD, on_click=go_next),
            ]
            self.page.update()

        render()
        return self._step_layout(
            "Step 4: Match columns",
            "",
            [holder],
            back_to=STEP_GRID,
        )

    def _build_job_and_plan(self):
        idg = self._grid_col(self.id_portal)
        mappings = []
        for pair in self.grade_pairs:
            if pair["portal"] is None or not pair["excel"]:
                continue
            g = self._grid_col(pair["portal"])
            mappings.append(automation.ColumnMapping(g.index, g.label, pair["excel"]))

        self.job = automation.FillJob(
            excel_path=self.excel_path,
            sheet_name=self.sheet_name,
            id_mapping=automation.ColumnMapping(idg.index, idg.label, self.id_excel),
            grade_mappings=mappings,
            blank_as_zero=self.blank_as_zero,
            round_up=self._round_up_rule()[0],
        )
        self.plan = automation.build_fill_plan(self.job, self.df, portal_ids=idg.values)

    # ---------- Step 5: Review, fill & save ----------

    def _build_step_run(self) -> ft.Control:
        job, plan = self.job, self.plan

        summary = [
            self._fact(
                ft.Icons.EDIT_NOTE,
                f"{len(plan.students)} students will be updated "
                f"({', '.join(m.portal_label for m in job.grade_mappings)}).",
                OK if plan.students else ERROR,
            )
        ]

        # Only things worth a second look stay visible; the ID lists and
        # minor notes go behind "Show details".
        warnings, details = [], []
        if plan.not_in_portal:
            warnings.append(f"{len(plan.not_in_portal)} students in Excel aren't in the portal table")
            details.append(f"In Excel but not in the portal (skipped): {_preview(plan.not_in_portal, 30)}")
        if plan.missing_from_excel:
            warnings.append(f"{len(plan.missing_from_excel)} portal students aren't in Excel")
            details.append(f"In the portal but not in Excel (left as they are): {_preview(plan.missing_from_excel, 30)}")
        if plan.duplicate_ids:
            warnings.append(f"{len(plan.duplicate_ids)} IDs appear twice in Excel")
            details.append(f"Appear more than once in Excel (skipped): {_preview(plan.duplicate_ids, 30)}")
        if plan.invalid_cells:
            warnings.append(f"{len(plan.invalid_cells)} cells aren't numbers")
            details.append("Not numbers (skipped):\n" + "\n".join(f"  • {c}" for c in plan.invalid_cells[:30]))
        if job.round_up is not None:
            n = len(plan.rounded)
            summary.append(self._fact(
                ft.Icons.TRENDING_UP,
                f"{n} grade{'s' if n != 1 else ''} will be rounded up ({self._round_summary(job.round_up)})."
                if n else "Round-up is on, but no grades needed rounding.",
                INFO,
            ))
            if n:
                details.append("Rounded up:\n" + "\n".join(f"  • {r}" for r in plan.rounded[:30])
                               + (f"\n  … and {n - 30} more" if n > 30 else ""))
        if plan.blank_cells:
            details.append(f"{plan.blank_cells} blank Excel cells will be left as they are.")
        if plan.rows_without_id:
            details.append(f"{plan.rows_without_id} Excel rows have no student ID (ignored).")

        if warnings:
            summary.append(self._fact(ft.Icons.WARNING_AMBER, " · ".join(warnings) + ".", WARN))
        if details:
            details_box = ft.Column(
                [ft.Text(d, size=12, selectable=True) for d in details], spacing=6, visible=False
            )

            def toggle_details(e):
                details_box.visible = not details_box.visible
                e.control.content = "Hide details" if details_box.visible else "Show details"
                self.page.update()

            summary += [ft.TextButton("Show details", on_click=toggle_details), details_box]

        preview_rows = plan.students[:5]
        preview = ft.DataTable(
            columns=[ft.DataColumn(ft.Text("Student ID", weight=ft.FontWeight.BOLD))]
            + [ft.DataColumn(ft.Text(m.portal_label, weight=ft.FontWeight.BOLD)) for m in job.grade_mappings],
            rows=[
                ft.DataRow(
                    cells=[ft.DataCell(ft.Text(s.sid))]
                    + [
                        ft.DataCell(
                            ft.Text(
                                next((str(w.value) for w in s.writes if w.mapping.portal_index == m.portal_index),
                                     "— unchanged")
                            )
                        )
                        for m in job.grade_mappings
                    ]
                )
                for s in preview_rows
            ],
            heading_row_height=36,
            data_row_min_height=30,
            data_row_max_height=30,
        )

        progress = ft.ProgressBar(value=0)
        progress_text = ft.Text("")
        log = ft.ListView(height=200, spacing=2, auto_scroll=True)
        result_box = ft.Column([], spacing=6)
        start_btn = ft.FilledButton("Start filling", icon=ft.Icons.PLAY_ARROW, disabled=not plan.students)
        stop_btn = ft.OutlinedButton("Stop", icon=ft.Icons.STOP, disabled=True)
        back_btn = ft.TextButton("Back to column matching", icon=ft.Icons.ARROW_BACK,
                                 on_click=lambda _: self._go_to(STEP_MAP))
        finish_btn = ft.FilledButton(
            "I clicked Save in the portal — close the browser",
            icon=ft.Icons.CHECK,
            visible=False,
        )

        def log_line(text: str, color=None):
            log.controls.append(ft.Text(text, size=12, font_family="monospace", color=color, selectable=True))

        def set_running(running: bool):
            self._running = running
            start_btn.disabled = running
            stop_btn.disabled = not running
            back_btn.disabled = running

        def on_start(_):
            self._stop_requested = False
            set_running(True)
            result_box.controls = []
            log.controls = []
            progress.value = 0
            progress_text.value = "Starting…"
            self.page.update()

            def on_progress(i, total, message):
                # Called on the browser thread; the screen update itself is
                # handed to Flet's loop (see _on_ui_thread), or the progress
                # bar would only redraw when the window lost focus.
                self._on_ui_thread(show_progress, i, total, message)

            def show_progress(i, total, message):
                progress.value = i / total if total else 0
                progress_text.value = f"{i} of {total}"
                log_line(message, WARN if message.startswith(("Not found", "Problem")) else None)
                self.page.update()

            def do_run():
                return automation.run_fill(
                    self._session(), job, plan,
                    on_progress=on_progress,
                    should_stop=lambda: self._stop_requested,
                )

            def on_done(result: automation.FillResult):
                set_running(False)
                show_result(result)
                if result.stopped:
                    progress_text.value = "Stopped. Press Start to run again from the beginning."
                else:
                    progress_text.value = "Finished typing the grades."
                    start_btn.content = "Fill again"
                finish_btn.visible = True
                self.page.update()

            def on_error(e):
                set_running(False)
                progress_text.value = "Stopped because of an error."
                result_box.controls = [ft.Text(self._error_text(e), color=ERROR)]
                log_line("".join(traceback.format_exception(e)), ERROR)
                self.page.update()

            self.pw.submit(do_run, on_done, on_error)

        def show_result(result: automation.FillResult):
            items = [self._fact(ft.Icons.CHECK_CIRCLE, f"{result.updated} students updated.", OK)]
            if result.not_found:
                items.append(self._fact(ft.Icons.WARNING_AMBER,
                                        f"Not found in the table: {_preview(result.not_found)}", WARN))
            if result.failed:
                items.append(self._fact(ft.Icons.ERROR_OUTLINE,
                                        f"{len(result.failed)} cells couldn't be set — check them by hand:", ERROR))
                items += [ft.Text(f"   • {f}", size=12, selectable=True) for f in result.failed[:20]]
                if len(result.failed) > 20:
                    items.append(ft.Text(f"   … and {len(result.failed) - 20} more (see log).", size=12))
            items.append(
                self._notice(
                    "Nothing is saved yet. Check the grades in the browser, then click the portal's own "
                    "Save button. Need to fill another column in this table? Go back to column matching."
                )
            )
            result_box.controls = items

        def on_stop(_):
            self._stop_requested = True
            stop_btn.disabled = True
            progress_text.value = "Stopping after the current student…"
            self.page.update()

        def on_finish(_):
            finish_btn.disabled = True
            self.page.update()

            def save_and_close():
                session = self._session()
                try:
                    session.save_session()
                finally:
                    session.close()

            def done(_r):
                self._show_message("All done", "The browser was closed. You can close this window now.")

            def failed(e):
                self._show_message("All done", "The browser was closed. You can close this window now.")

            self.pw.submit(save_and_close, done, failed)

        start_btn.on_click = on_start
        stop_btn.on_click = on_stop
        finish_btn.on_click = on_finish

        return self._step_layout(
            "Step 5: Review and fill",
            "Check the summary below. Filling types the grades into the portal table; "
            "nothing is saved until you click Save in the portal yourself.",
            [
                self._card("What will happen", summary),
                self._card(
                    f"Preview (first {len(preview_rows)} of {len(plan.students)} students)",
                    [ft.Row([preview], scroll=ft.ScrollMode.AUTO)],
                ) if preview_rows else ft.Container(),
                ft.Row([start_btn, stop_btn], wrap=True),
                progress,
                progress_text,
                ft.Container(content=log, bgcolor=ft.Colors.BLUE_GREY_50, padding=8, border_radius=8),
                result_box,
                finish_btn,
                back_btn,
            ],
        )

    # ---------- small UI helpers ----------

    @staticmethod
    def _card(title: str, controls: list) -> ft.Control:
        return ft.Container(
            content=ft.Column(
                [ft.Text(title, weight=ft.FontWeight.BOLD, size=15)] + controls,
                spacing=8,
            ),
            padding=16,
            border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
            border_radius=10,
        )

    @staticmethod
    def _notice(text: str) -> ft.Control:
        return ft.Container(
            content=ft.Row(
                [ft.Icon(ft.Icons.INFO_OUTLINE, color=ft.Colors.BLUE_700), ft.Text(text, expand=True, size=13)],
                vertical_alignment=ft.CrossAxisAlignment.START,
            ),
            bgcolor=ft.Colors.BLUE_50,
            padding=12,
            border_radius=8,
        )

    @staticmethod
    def _fact(icon, text: str, color=INFO, action: Optional[ft.Control] = None) -> ft.Control:
        if action is None:
            # Long text must wrap inside the row, so let it take the remaining width.
            return ft.Row(
                [ft.Icon(icon, color=color, size=18), ft.Text(text, expand=True, selectable=True)],
                vertical_alignment=ft.CrossAxisAlignment.START,
            )
        return ft.Row(
            [ft.Icon(icon, color=color, size=18), ft.Text(text, selectable=True), action],
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            wrap=True,
        )

    def _show_message(self, title: str, message: str):
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text(title),
            content=ft.Text(message),
            actions=[ft.TextButton("OK", on_click=lambda _: self.page.pop_dialog())],
        )
        self.page.show_dialog(dlg)

    # ---------- shutdown ----------

    def _on_window_event(self, e):
        if e.type != ft.WindowEventType.CLOSE:
            return
        self._stop_requested = True
        closed = threading.Event()

        def destroy(*_):
            if not closed.is_set():
                closed.set()
                self.page.run_task(self.page.window.destroy)

        def close_browser():
            if self.session:
                self.session.close()

        self.pw.submit(close_browser, destroy, destroy)
        # Don't hang forever if the worker is stuck mid-action.
        threading.Timer(6.0, destroy).start()


def main(page: ft.Page):
    setup_logging()
    GradeFillerApp(page)


if __name__ == "__main__":
    ft.run(main)
