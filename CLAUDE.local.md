# Grade Filler: where we left off

> **For every agent working in this repo:** read this file before changing
> anything, and **keep it updated** before you finish a task. Do this in the same
> session as the work:
> - move items from "Open issues" / "Next features" to "Recent fixes" when done
> - add any new bug, limitation or user decision you discover
> - update "Last updated" and the branch/push status
> - keep entries short: what, why, where (file/function), commit hash if any
>
> This file is committed to the repo on purpose, so it's backed up on GitHub and
> comes with every download. (The user lost an earlier set of notes once.) Commit
> your updates to it along with your code changes.

**Last updated:** 2026-11-19 · **Branch:** `improving-UI` (code pushed; these notes committed, not yet pushed) · **`main`:** old version, not merged yet

## What this is

A desktop app (Python + Flet UI + Playwright) that types grades from a teacher's
Excel sheet into the Al-Quds University grades portal
(`portal.alquds.edu`, a DHTMLX grid). Nothing is saved automatically. The
teacher reviews and clicks the portal's own Save button.

**The real user** runs it on a **Windows laptop and isn't technical**: no git,
terminals are a struggle. Every UI text must be plain language. The end goal is a
double-click `.exe` with no terminal at all.

## Layout

| File | Role |
|---|---|
| `main.py` | Launcher: `python main.py` |
| `gradefiller/app.py` | Flet wizard, 5 steps: Log In → Grades File → Course Table → Match Columns → Review & Fill |
| `gradefiller/automation.py` | All browser + Excel work: `PortalSession`, grid discovery (`find_grids`, `GridInfo`), `scan_grid`, `load_sheet` / `excel_display_text`, `build_fill_plan`, `run_fill`, `edit_cell_commit` |
| `gradefiller/matching.py` | Pure auto-matching: ID pair by shared values, grade pairs by name similarity |
| `gradefiller/playwright_worker.py` | One thread owns Playwright (sync API is thread-bound); UI submits tasks |
| `gradefiller/paths.py` | Saved login → `%APPDATA%\GradeFiller\state.json` / `~/Library/Application Support/GradeFiller` |
| `GUI.py`, `upload_grades.py`, `save_session.py` | **Old, obsolete** (Tkinter / hardcoded to one course). Unused. Delete once the user agrees. |

## Run it

```
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt     # macOS: .venv/bin/python
.venv\Scripts\python -m playwright install chromium
.venv\Scripts\python main.py
```
On Windows, skip `activate` (script execution policy errors). Call `.venv\Scripts\python` directly.
New computer without git: download https://github.com/Hala036/Grade-Filler/archive/refs/heads/improving-UI.zip

## Decisions the user made (don't undo without asking)

- **Works for any course / any number of exams.** Nothing hardcoded; columns are read from the live grid and mapped in the UI.
- **Grades are copied exactly as Excel displays them** (number formats applied: format `0` showing 14 types 14, General 12.5 types 12.5, float noise removed). No rounding beyond what Excel shows. (`excel_display_text`)
- **Blank Excel cells leave the portal cell unchanged** by default. A checkbox writes 0 instead. (I chose this; the user didn't object.)
- **Duplicate IDs in Excel are skipped**, never guessed. Non-numeric cells are skipped and reported.

## Recent fixes (newest first)

- `7bc480c`: Copy grades as Excel shows them (they used to be rounded half-up to whole numbers).
- `91fcf04`: **Popup grade table.** The portal lists groups in one grid and opens the real grade table in a popup. The old scan mixed all grids on the page and read the group list. Now every grid (all tabs and iframes) is discovered and tagged with `data-gradefiller-grid`; the visible one in front wins; Match Columns has a "Table to fill" picker when there are several. The edit focus click moved from page centre to the row's ID cell.
- `85f61ab`: `requirements.txt` pins only direct deps (old `numpy==2.0.2` had no Python 3.13+ wheels, so the Windows install failed).
- `648351b`: Wizard rework. Flet 0.86 fixes (async `FilePicker.pick_files`, `Dropdown.on_select`, button `content`); one reused browser instead of two; retry/Back everywhere; worker survives callback errors; auto-matching; example values in dropdowns; validation; review step with preview; per-cell failures don't abort; end-of-run summary.

## Open issues / unverified (check on the REAL portal)

- [ ] **Popup detection** was tested only on fake pages (same-DOM modal, iframe, new window). Confirm on the real portal; if it's wrong, inspect the popup's DOM and adjust `_DISCOVER_JS` / `GridInfo.priority`.
- [ ] **Edit focus click on the ID cell** (replaced the old `body.click()`). Confirm edits still commit, and that clicking a row does nothing unwanted.
- [ ] **Decimals:** does the portal accept `36.8` / `25.0`? If it only takes integers, those cells show as "couldn't be set".
- [ ] **Header labels:** confirm portal column names are read. Otherwise the UI falls back to "Column N" with example values.
- [ ] Merge `improving-UI` → `main` only after a successful real run.
- [ ] Delete the obsolete `GUI.py`, `upload_grades.py`, `save_session.py` (ask first).
- [ ] No automated tests in the repo yet. Earlier checks were throwaway scripts: fake DHTMLX pages + headless Chromium, and a stub Flet page driving each step. Worth adding under `tests/`.
- [ ] Minor: the portal's row-number column (`#`) is offered as a fill target. When names tie, suggestions pick the left-most (`First/30` vs `First/25` in sec1).

## Next features

1. **Windows `.exe`, no terminal (top priority).** Plan: GitHub Actions on `windows-latest` builds a single file (`flet pack`, which wraps PyInstaller) from `main.py`, uploaded as a build artifact/release. Saved login already lives in `%APPDATA%`. Expect a SmartScreen "More info → Run anyway" on first launch (unsigned). Needs the user's go-ahead before pushing a workflow.
   - **Don't bundle Chromium; use the installed Edge.** Today the app drives Playwright's own downloaded Chromium (`playwright install chromium`, ~150 MB in a user cache folder). Packing that into the `.exe` makes it huge, and Playwright's browser-path lookup breaks inside PyInstaller's temp extraction. Instead launch with `chromium.launch(channel="msedge")` (Edge ships with every Windows install), falling back to `channel="chrome"`, then to the downloaded Chromium for dev machines. Change is in `PortalSession._ensure_browser`. Small risk: a future auto-updated Edge could outrun the pinned Playwright, in which case bump `playwright` in `requirements.txt` and rebuild.
2. Whatever the user asks for next. Add it here.
