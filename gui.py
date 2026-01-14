import math
import sys
import threading
import time
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PwTimeoutError, Error as PwError

# ===== Portal settings =====
COURSE_PAGE_URL = "https://portal.alquds.edu/en/registration/marks/course_list"
ROW_LOCATOR = "tr.ev_dhx_web, tr.odd_dhx_web"

# Portal column indexes (0-based td)
IDX_STUDENT_NUMBER = 1

DEFAULT_STATE_PATH = "state.json"
PACE_MS = 120


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


def row_for_student(page, sid: str, id_td_index: int):
    rows = page.locator(ROW_LOCATOR)
    td_n = id_td_index + 1
    # exact match if supported
    try:
        return rows.filter(
            has=page.locator(f'td:nth-child({td_n}):text-is("{sid}")')
        ).first
    except Exception:
        return rows.filter(
            has=page.locator(f"td:nth-child({td_n})", has_text=sid)
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
        self.map_rows = []
        self._mappings = []

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

        f1.columnconfigure(1, weight=1)

        f2 = ttk.LabelFrame(self, text="2) Sheet")
        f2.pack(fill="x", **pad)

        ttk.Label(f2, text="Sheet:").grid(row=0, column=0, sticky="e", **pad)
        self.sheet_combo = ttk.Combobox(f2, textvariable=self.sheet_name, state="readonly", width=40)
        self.sheet_combo.grid(row=0, column=1, sticky="w", **pad)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self.load_columns())

        f3 = ttk.LabelFrame(self, text="3) Column mapping")
        f3.pack(fill="x", **pad)

        self.map_frame = f3
        ttk.Label(f3, text="Index (1-based):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Label(f3, text="Name:").grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(f3, text="Sheet column:").grid(row=0, column=2, sticky="w", **pad)

        self._add_mapping_row(name="Student ID", idx=IDX_STUDENT_NUMBER + 1)
        self.map_add_btn = ttk.Button(f3, text="Add column", command=self.add_mapping_row)
        self._refresh_mapping_controls()

        f4 = ttk.LabelFrame(self, text="4) Run")
        f4.pack(fill="x", **pad)

        self.btn_start = ttk.Button(f4, text="Start Fill (no save)", command=self.start)
        self.btn_start.grid(row=0, column=0, **pad)

        self.btn_stop = ttk.Button(f4, text="Stop", command=self.stop, state="disabled")
        self.btn_stop.grid(row=0, column=1, **pad)

        self.progress = ttk.Progressbar(f4, length=400, mode="determinate")
        self.progress.grid(row=1, column=0, columnspan=4, sticky="we", **pad)

        f5 = ttk.LabelFrame(self, text="Log")
        f5.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(f5, height=18, wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

    def _add_mapping_row(self, name="", idx=""):
        pad = {"padx": 10, "pady": 6}
        row = len(self.map_rows) + 1
        idx_var = tk.StringVar(value=str(idx) if idx else "")
        name_var = tk.StringVar(value=name)
        sheet_var = tk.StringVar()

        idx_entry = ttk.Entry(self.map_frame, textvariable=idx_var, width=4)
        name_entry = ttk.Entry(self.map_frame, textvariable=name_var, width=15)
        sheet_combo = ttk.Combobox(self.map_frame, textvariable=sheet_var, state="readonly", width=15)
        row_data = {
            "idx_var": idx_var,
            "name_var": name_var,
            "sheet_var": sheet_var,
            "sheet_combo": sheet_combo,
        }
        remove_btn = ttk.Button(self.map_frame, text="Remove", command=lambda r=row_data: self.remove_mapping_row(r))

        idx_entry.grid(row=row, column=0, sticky="w", **pad)
        name_entry.grid(row=row, column=1, sticky="w", **pad)
        sheet_combo.grid(row=row, column=2, sticky="w", **pad)
        remove_btn.grid(row=row, column=3, sticky="w", **pad)

        row_data["widgets"] = (idx_entry, name_entry, sheet_combo, remove_btn)
        self.map_rows.append(row_data)

    def _refresh_mapping_controls(self):
        pad = {"padx": 10, "pady": 6}
        for i, row in enumerate(self.map_rows, start=1):
            for w in row["widgets"]:
                w.grid_configure(row=i)
        self.map_add_btn.grid(row=len(self.map_rows) + 1, column=0, sticky="w", **pad)

    def add_mapping_row(self):
        self._add_mapping_row()
        if self.map_rows:
            cols = self.map_rows[0]["sheet_combo"]["values"]
            if cols:
                self.map_rows[-1]["sheet_combo"]["values"] = cols
        self._refresh_mapping_controls()

    def remove_mapping_row(self, row_data):
        if not self.map_rows or self.map_rows[0] is row_data:
            return
        try:
            self.map_rows.remove(row_data)
        except ValueError:
            return
        for w in row_data["widgets"]:
            w.destroy()
        self._refresh_mapping_controls()

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
            self.load_sheets()

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
                self.load_columns()
            self.log_line(f"[OK] Sheets loaded: {list(xls.sheet_names)}")
        except Exception as e:
            messagebox.showerror("Excel error", str(e))

    def load_columns(self):
        path = self.excel_path.get().strip()
        sheet_raw = self.sheet_name.get()
        if not path or not sheet_raw.strip():
            messagebox.showerror("Missing", "Pick Excel + sheet first.")
            return
        df = pd.read_excel(path, sheet_name=sheet_raw, nrows=2)
        cols = list(df.columns)
        for row in self.map_rows:
            row["sheet_combo"]["values"] = cols
        self.log_line(f"[OK] Columns loaded ({len(cols)}). Now map them.")

    def _collect_mappings(self):
        if not self.map_rows:
            messagebox.showerror("Missing", "Add at least one column mapping.")
            return None

        mappings = []
        for i, row in enumerate(self.map_rows, start=1):
            idx_raw = row["idx_var"].get().strip()
            name = row["name_var"].get().strip()
            sheet_col = row["sheet_var"].get().strip()

            if not idx_raw or not name or not sheet_col:
                messagebox.showerror("Missing", f"Complete mapping row {i}.")
                return None
            try:
                idx = int(idx_raw)
            except ValueError:
                messagebox.showerror("Invalid", f"Index must be a number in row {i}.")
                return None
            if idx < 1:
                messagebox.showerror("Invalid", f"Index must be 1 or greater in row {i}.")
                return None

            mappings.append(
                {"idx": idx - 1, "name": name, "sheet_col": sheet_col, "is_id": i == 1}
            )

        return mappings

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
        mappings = self._collect_mappings()
        if not mappings:
            return

        self._stop = False
        self._mappings = mappings
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
            sheet = self.sheet_name.get()
            state = self.state_path.get().strip()
            mappings = list(self._mappings)
            id_map = mappings[0]
            value_maps = mappings[1:]
            cols = [m["sheet_col"] for m in mappings]

            self.log_line("[INFO] Loading Excel…")
            df = pd.read_excel(excel, sheet_name=sheet)[cols].copy()

            df[id_map["sheet_col"]] = df[id_map["sheet_col"]].apply(normalize_id)
            for m in value_maps:
                df[m["sheet_col"]] = df[m["sheet_col"]].apply(normalize_grade)

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

                pace = max(0, int(PACE_MS)) / 1000.0

                browser_ok = True
                for i, r in enumerate(df.itertuples(index=False, name=None), start=1):
                    if self._stop:
                        self.log_line("[STOP] Stopped by user.")
                        break

                    row_data = dict(zip(cols, r))
                    sid = str(row_data[id_map["sheet_col"]])
                    row = row_for_student(page, sid, id_map["idx"])
                    try:
                        row.wait_for(timeout=6000)
                    except PwTimeoutError:
                        self.log_line(f"[MISSING] {sid}")
                        self.progress["value"] = i
                        continue
                    except PwError as e:
                        browser_ok = False
                        self.log_line(f"[ERROR] {e}")
                        self._popup("Browser closed", "The browser closed unexpectedly. Click Start again.")
                        break

                    for m in value_maps:
                        edit_cell_commit(page, row, m["idx"], int(row_data[m["sheet_col"]]))

                    if i % 10 == 0:
                        self.log_line(f"[PROGRESS] {i}/{total}")

                    self.progress["value"] = i
                    if pace:
                        time.sleep(pace)

                if browser_ok:
                    self.log_line("[DONE] Filling completed.")
                    self._popup("Save", "Now click Save in the portal and confirm the popup.\nClick OK here when you are done, to close the browser.")

                    context.storage_state(path=state)
                    browser.close()

        except Exception as e:
            self.log_line(f"[ERROR] {e}")
            self.log_line(traceback.format_exc().strip())
            if "Event loop is closed" in str(e):
                self._popup("Browser closed", "The browser closed unexpectedly. Click Start again.")
            else:
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
