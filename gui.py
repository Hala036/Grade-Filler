import math
import sys
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PwTimeoutError

# ===== Portal settings =====
COURSE_PAGE_URL = "https://portal.alquds.edu/en/registration/marks/course_list"
ROW_LOCATOR = "tr.ev_dhx_web, tr.odd_dhx_web"

# Portal column indexes (0-based td)
IDX_STUDENT_NUMBER = 1
IDX_FIRST = 3
IDX_SECOND = 4
IDX_QUIZ = 5
IDX_FINAL = 6

DEFAULT_STATE_PATH = "state.json"


# ===== helpers =====
def round_half_up(x: float) -> int:
    return int(math.floor(x + 0.5))


def normalize_id(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s == "":
        return ""
    return str(int(float(s)))


def normalize_grade(x) -> int:
    if pd.isna(x) or str(x).strip() == "":
        return 0
    return round_half_up(float(x))


def row_for_student(page, sid: str):
    rows = page.locator(ROW_LOCATOR)
    # exact match if supported
    try:
        return rows.filter(
            has=page.locator(f'td:nth-child({IDX_STUDENT_NUMBER + 1}):text-is("{sid}")')
        ).first
    except Exception:
        return rows.filter(
            has=page.locator(f"td:nth-child({IDX_STUDENT_NUMBER + 1})", has_text=sid)
        ).first


def edit_cell_commit(page, row, td_index: int, value: int):
    """Cross-platform safe edit: Cmd+A on mac (Meta), Ctrl+A on windows."""
    modifier = "Meta" if sys.platform == "darwin" else "Control"

    page.bring_to_front()
    page.locator("body").click()

    cell = row.locator("td").nth(td_index)
    cell.scroll_into_view_if_needed()

    # DHTMLX can be finicky: retry a couple times
    for _ in range(3):
        cell.dblclick()
        page.wait_for_timeout(80)

        page.keyboard.press(f"{modifier}+A")
        page.keyboard.insert_text(str(int(value)))
        page.keyboard.press("Tab")
        page.wait_for_timeout(150)

        # if it shows something, good enough
        if cell.inner_text().strip() != "":
            return

        page.wait_for_timeout(200)

    raise RuntimeError(f"Edit failed: td={td_index} value={value}")


# ===== GUI app =====
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Grade Filler (Fill only)")
        self.geometry("900x640")

        self.excel_path = tk.StringVar()
        self.state_path = tk.StringVar(value=str(Path.cwd() / DEFAULT_STATE_PATH))
        self.sheet_name = tk.StringVar()

        self.col_id = tk.StringVar()
        self.col_first = tk.StringVar()
        self.col_second = tk.StringVar()
        self.col_quiz = tk.StringVar()
        self.col_final = tk.StringVar()

        self.pace_ms = tk.IntVar(value=120)

        self._stop = False
        self._worker = None

        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 6}

        f1 = ttk.LabelFrame(self, text="1) Files")
        f1.pack(fill="x", **pad)

        ttk.Label(f1, text="Excel:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.excel_path, width=70).grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(f1, text="Browse…", command=self.pick_excel).grid(row=0, column=2, **pad)

        ttk.Label(f1, text="state.json:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.state_path, width=70).grid(row=1, column=1, sticky="we", **pad)
        ttk.Button(f1, text="Browse…", command=self.pick_state).grid(row=1, column=2, **pad)

        f1.columnconfigure(1, weight=1)

        f2 = ttk.LabelFrame(self, text="2) Sheet")
        f2.pack(fill="x", **pad)

        ttk.Button(f2, text="Load sheets", command=self.load_sheets).grid(row=0, column=0, **pad)
        ttk.Label(f2, text="Sheet:").grid(row=0, column=1, sticky="e", **pad)
        self.sheet_combo = ttk.Combobox(f2, textvariable=self.sheet_name, state="readonly", width=40)
        self.sheet_combo.grid(row=0, column=2, sticky="w", **pad)

        f3 = ttk.LabelFrame(self, text="3) Column mapping")
        f3.pack(fill="x", **pad)

        self.cols_combo = {}
        self._map_row(f3, 0, "Student ID", self.col_id)
        self._map_row(f3, 1, "First", self.col_first)
        self._map_row(f3, 2, "Second", self.col_second)
        self._map_row(f3, 3, "Quiz", self.col_quiz)
        self._map_row(f3, 4, "Final", self.col_final)

        ttk.Button(f3, text="Load columns", command=self.load_columns).grid(row=5, column=0, **pad)
        ttk.Button(f3, text="Auto-guess", command=self.auto_guess).grid(row=5, column=1, **pad)

        f4 = ttk.LabelFrame(self, text="4) Run")
        f4.pack(fill="x", **pad)

        ttk.Label(f4, text="Pace (ms):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Spinbox(f4, from_=0, to=2000, textvariable=self.pace_ms, width=8).grid(row=0, column=1, sticky="w", **pad)

        self.btn_start = ttk.Button(f4, text="Start Fill (no save)", command=self.start)
        self.btn_start.grid(row=0, column=2, **pad)

        self.btn_stop = ttk.Button(f4, text="Stop", command=self.stop, state="disabled")
        self.btn_stop.grid(row=0, column=3, **pad)

        self.progress = ttk.Progressbar(f4, length=400, mode="determinate")
        self.progress.grid(row=1, column=0, columnspan=4, sticky="we", **pad)

        f5 = ttk.LabelFrame(self, text="Log")
        f5.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(f5, height=18, wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

    def _map_row(self, parent, r, label, var):
        pad = {"padx": 10, "pady": 6}
        ttk.Label(parent, text=f"{label}:").grid(row=r, column=0, sticky="w", **pad)
        cb = ttk.Combobox(parent, textvariable=var, state="readonly", width=45)
        cb.grid(row=r, column=1, sticky="w", **pad)
        self.cols_combo[label] = cb

    def log_line(self, s):
        self.log.insert("end", s + "\n")
        self.log.see("end")
        self.update_idletasks()

    def pick_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)
            self.sheet_name.set("")
            self.log_line(f"[OK] Excel selected: {path}")

    def pick_state(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            self.state_path.set(path)

    def load_sheets(self):
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showerror("Missing", "Pick an Excel file first.")
            return
        try:
            xls = pd.ExcelFile(path)
            self.sheet_combo["values"] = list(xls.sheet_names)
            if xls.sheet_names:
                self.sheet_name.set(xls.sheet_names[0])
            self.log_line(f"[OK] Sheets loaded: {list(xls.sheet_names)}")
        except Exception as e:
            messagebox.showerror("Excel error", str(e))

    def load_columns(self):
        path = self.excel_path.get().strip()
        sheet = self.sheet_name.get().strip()
        if not path or not sheet:
            messagebox.showerror("Missing", "Pick Excel + sheet first.")
            return
        df = pd.read_excel(path, sheet_name=sheet, nrows=2)
        cols = list(df.columns)
        for cb in self.cols_combo.values():
            cb["values"] = cols
        self.log_line(f"[OK] Columns loaded ({len(cols)}). Now map them.")

    def auto_guess(self):
        cols = self.cols_combo["Student ID"]["values"]
        if not cols:
            messagebox.showerror("Missing", "Click 'Load columns' first.")
            return

        def pick(*needles):
            for c in cols:
                s = str(c).lower()
                if all(n in s for n in needles):
                    return c
            return ""

        self.col_id.set(pick("id") or pick("id number"))
        self.col_first.set(pick("first", "/25") or pick("first"))
        self.col_second.set(pick("second", "/25") or pick("second"))
        self.col_quiz.set(pick("quiz") or pick("qiuz"))
        self.col_final.set(pick("final") or pick("grade", "40") or pick("40"))

        self.log_line("[OK] Auto-guess filled. Please verify mappings.")

    def start(self):
        if self._worker and self._worker.is_alive():
            return

        # basic checks
        if not self.excel_path.get().strip():
            messagebox.showerror("Missing", "Pick an Excel file.")
            return
        if not self.sheet_name.get().strip():
            messagebox.showerror("Missing", "Pick a sheet.")
            return
        if not Path(self.state_path.get().strip()).exists():
            messagebox.showerror("Missing", "state.json not found. Create it first.")
            return
        for name, var in [("Student ID", self.col_id), ("First", self.col_first), ("Second", self.col_second), ("Quiz", self.col_quiz), ("Final", self.col_final)]:
            if not var.get().strip():
                messagebox.showerror("Missing", f"Map column for: {name}")
                return

        self._stop = False
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress["value"] = 0

        self._worker = threading.Thread(target=self._run_worker, daemon=True)
        self._worker.start()

    def stop(self):
        self._stop = True
        self.log_line("[STOP] Stop requested. Will finish current student then stop.")

    def _run_worker(self):
        try:
            excel = self.excel_path.get().strip()
            sheet = self.sheet_name.get().strip()
            state = self.state_path.get().strip()

            col_id = self.col_id.get().strip()
            col_first = self.col_first.get().strip()
            col_second = self.col_second.get().strip()
            col_quiz = self.col_quiz.get().strip()
            col_final = self.col_final.get().strip()

            self.log_line("[INFO] Loading Excel…")
            df = pd.read_excel(excel, sheet_name=sheet)[[col_id, col_first, col_second, col_quiz, col_final]].copy()

            df[col_id] = df[col_id].apply(normalize_id)
            for c in [col_first, col_second, col_quiz, col_final]:
                df[c] = df[c].apply(normalize_grade)

            total = len(df)
            self.progress["maximum"] = total
            self.log_line(f"[OK] Loaded {total} students.")

            # Instruction
            self._popup("Navigate", "A browser will open.\nLog in if needed, go to the grades table page,\nthen click OK here to start filling.")

            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False)
                context = browser.new_context(storage_state=state)
                page = context.new_page()
                page.goto(COURSE_PAGE_URL, wait_until="domcontentloaded")

                # Wait for user to navigate
                self._popup("Ready?", "Now in the browser, navigate to the correct course/section grade table.\nWhen you see the student grid, click OK here.")

                # Ensure grid exists
                page.locator(ROW_LOCATOR).first.wait_for(timeout=60_000)
                self.log_line(f"[OK] Grid detected. Rows: {page.locator(ROW_LOCATOR).count()}")

                pace = max(0, int(self.pace_ms.get())) / 1000.0

                for i, r in enumerate(df.itertuples(index=False), start=1):
                    if self._stop:
                        self.log_line("[STOP] Stopped by user.")
                        break

                    sid = getattr(r, col_id.replace(" ", "_")) if False else str(r[0])  # ignore; we use index positions below

                    sid = str(r[0])
                    first = int(r[1]); second = int(r[2]); quiz = int(r[3]); final = int(r[4])

                    row = row_for_student(page, sid)
                    try:
                        row.wait_for(timeout=6000)
                    except PwTimeoutError:
                        self.log_line(f"[MISSING] {sid}")
                        self.progress["value"] = i
                        continue

                    edit_cell_commit(page, row, IDX_FIRST, first)
                    edit_cell_commit(page, row, IDX_SECOND, second)
                    edit_cell_commit(page, row, IDX_QUIZ, quiz)
                    edit_cell_commit(page, row, IDX_FINAL, final)

                    if i % 10 == 0:
                        self.log_line(f"[PROGRESS] {i}/{total}")

                    self.progress["value"] = i
                    if pace:
                        time.sleep(pace)

                self.log_line("[DONE] Filling completed.")
                self._popup("Save", "Now click Save in the portal and confirm the popup.\nClick OK here when you are done, to close the browser.")

                context.storage_state(path=state)
                browser.close()

        except Exception as e:
            self.log_line(f"[ERROR] {e}")
            self._popup("Error", str(e))
        finally:
            self.btn_start.config(state="normal")
            self.btn_stop.config(state="disabled")

    def _popup(self, title, msg):
        ev = threading.Event()
        def _show():
            messagebox.showinfo(title, msg)
            ev.set()
        self.after(0, _show)
        ev.wait()


if __name__ == "__main__":
    App().mainloop()
