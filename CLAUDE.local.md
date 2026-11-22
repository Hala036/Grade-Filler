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

**Last updated:** 2026-11-22 · **Released: v1.1.0** (update notice, no login step + saved password, Match Columns options; PRs #3, #4, #5 merged into `main`) · before that v1.0.0 · **`main`:** current, working version · **Work happens on one branch per feature**, merged into `main` by pull request

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
| `main.py` | Launcher: `python main.py`; `--self-test` checks a build without opening the app |
| `assets/icon.png`, `assets/icon.ico` | App icon (user's red circled "A+"). `.ico` holds 16–256 px; 16–48 px have thickened strokes so they stay readable. Used for the `.exe` (`--icon`) and the window (`page.window.icon`, Windows only), bundled via `--add-data` and found with `paths.resource_path` |
| `gradefiller/selftest.py` | Self-test: imports, Playwright driver, headless Edge/Chrome; writes `selftest.txt`, exit 0 = pass |
| `.github/workflows/build-windows.yml` | Builds `GradeFiller.exe` on GitHub (Windows): manual run → Artifacts; tag `v*` → Release |
| `gradefiller/app.py` | Flet wizard, 4 steps: Grades File → Course Table → Match Columns → Review & Fill; account menu top-right |
| `gradefiller/automation.py` | All browser + Excel work: `PortalSession`, grid discovery (`find_grids`, `GridInfo`), `scan_grid`, `load_sheet` / `excel_display_text`, `build_fill_plan`, `run_fill`, `edit_cell_commit` |
| `gradefiller/matching.py` | Pure auto-matching: ID pair by shared values, grade pairs by name similarity |
| `gradefiller/playwright_worker.py` | One thread owns Playwright (sync API is thread-bound); UI submits tasks |
| `gradefiller/settings.py` | Small preferences between runs (`settings.json` in the app data folder); never secrets. So far: "don't ask to save my password" |
| `gradefiller/credentials.py` | The saved portal login (one account) in Windows Credential Manager / macOS Keychain via `keyring`. Never in a file or the log |
| `gradefiller/updater.py` | Update notice + self-update: asks GitHub's `releases/latest` API, downloads the new `.exe` (size + SHA-256 checked), swaps it in, restarts. No UI code |
| `gradefiller/_version.py` | The running version. `"0.0.0"` in git; the build overwrites it with the tag's number |
| `tests/` | `python -m unittest discover tests` (built-in, no extra installs). `test_updater.py`; `test_auto_login.py` (fake portal + fake SSO page in a hidden browser) |
| `gradefiller/paths.py` | Saved login → `%APPDATA%\GradeFiller\state.json` / `~/Library/Application Support/GradeFiller` |

## Run it

```
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt     # macOS: .venv/bin/python
.venv\Scripts\python -m playwright install chromium   # optional: only if neither Edge nor Chrome is installed
.venv\Scripts\python main.py
```
On Windows, skip `activate` (script execution policy errors). Call `.venv\Scripts\python` directly.
New computer without git: download https://github.com/Hala036/Grade-Filler/archive/refs/heads/main.zip

**The teacher's download link** (always the newest release): https://github.com/Hala036/Grade-Filler/releases/latest/download/GradeFiller.exe

**Releasing a new version:** merge the PR into `main`, then `git tag -a vX.Y.Z -m "…" && git push origin vX.Y.Z`. The workflow builds, self-tests and publishes the Release. Numbers: 1.0.1 = small fix, 1.1.0 = new feature, 2.0.0 = big change. Test builds (PRs, manual runs) are labelled 0.0.0 on purpose.

## Decisions the user made (don't undo without asking)

- **Works for any course / any number of exams.** Nothing hardcoded; columns are read from the live grid and mapped in the UI.
- **Grades are copied exactly as Excel displays them** (number formats applied: format `0` showing 14 types 14, General 12.5 types 12.5, float noise removed). No rounding beyond what Excel shows. (`excel_display_text`)
- **Blank Excel cells leave the portal cell unchanged** (always; the "write 0 for blank cells" checkbox was removed, user 2026-11-22).
- **A 0 in Excel leaves the portal cell unchanged too**, by default: checkbox "If a grade is 0, leave the portal cell as it is", ticked by default (user 2026-11-22). Unticked, 0 is typed. It looks at the Excel value, not at grades that become 0 after rounding. If the portal already holds a number, it stays (user's choice over erasing it). `FillJob.zero_as_blank`, `FillPlan.zero_cells`.
- **Duplicate IDs in Excel are skipped**, never guessed. Non-numeric cells are skipped and reported.
- **No login step; the browser opens only when the teacher clicks "Open the portal" in Course Table**, never by itself on arriving there (user, 2026-11-22). If the portal asks to log in, the saved password is used, or the teacher logs in there. **No background cookie check during Step 1** (discussed 2026-11-22: it only saves a few seconds and would need its own HTTP login check).
- **Saved password: the teacher logs in on the university page, then the app asks "Save your portal password?"** (not an in-app form). **One saved account**; the top-right menu has "Switch account" (logs out + forgets) and "Forget my password". Both chosen by the user, 2026-11-22. **The save question must appear after every login the teacher does by hand** (user, 2026-11-22), with "Don't ask again" (kept in `settings.json`; undone from the corner menu: "Offer to save my password again"). "Not now" only skips this one time.
- **Optional rounding (off by default)** is the one exception to "no rounding". Points snap to **0.05 steps** and the whole decimal part is compared (not just the first digit). The teacher's usual rule: .1–.49 up to .5, .5–.99 up to the next whole number (12.2 → 12.5, 12.5 → 13, 12 stays 12). **Decimals below the first arrow drop to .0** (12.05 → 12; user 2026-11-22, which replaces the old "rounding must never lower a grade"); decimals past the last arrow's end stay.

## Recent fixes (newest first)

- **Match Columns options** (branch `match-columns-options`, built on `auto-login`, 2026-11-22). Rounding line in 0.05 steps: `automation.ROUND_STEPS = 20`, `RoundUp(start, joint, end)` now in twentieths (default `(2, 10, 20)`), `RoundUp.target(frac)` takes the whole decimal part; `step_label` / `step_value` replace `round_line.tenth_label`. The line has numbers at tenths only, short ticks at the half-steps, and a small bold label above a dot sitting on a half-step. Checkbox renamed "Round grades". **Texts kept minimal (user, 2026-11-22: "as simple as possible; it's obvious")**: under the line only the rule, e.g. "below 0.25 → 0 · 0.25–0.49 → 0.5 · 0.5–0.89 → 0.9 · 0.9 and above stays". Zero option: see Decisions. Tests: `tests/test_rounding.py`. Drawing checked with `ft.Screenshot`; dragging on half-steps not tried by hand yet.

- **No login step + saved password / auto-login** (branch `auto-login`, 2026-11-22). Steps are now Grades File → Course Table → Match Columns → Review & Fill. "Open the portal" in Course Table opens the course list (nothing opens on arriving). After a saved-password login the fresh cookies are written straight away (`save_session`). If the login page is still showing, the browser task waits (`wait_for_manual_login`, up to 10 min; ends early if the browser is closed, on Switch account or on closing the app via `stop_waiting`), then saves the cookies, goes to the course list and the app asks to save the password **right then** (first version only asked when the teacher clicked "The grade table is showing", so the user, testing on the Mac, never saw it). The portal's login is the university SSO at `login.alquds.edu` (SimpleSAMLphp: `#username`, `#password`, `#submit_button`, no CAPTCHA seen; checked 2026-11-22). `automation._LOGIN_PAGE_JS` (an init script on every page, acts only on `SSO_HOST`): (1) reports what the teacher typed via the `gfLoginTyped` binding → `PortalSession.typed_login`, memory only; (2) if a login is saved (`window.__gfLogin`, set by `_LOGIN_CONFIG_JS`; the newest wins), fills it in and presses Log in, **at most once per 2 min per tab** (sessionStorage), so a wrong saved password never loops; a page that kept the username = failed login → `window.__gfLoginFailed`, and `open_course_list` stops waiting at once. Both the open task and `_scan` call `app._hand_over_typed_login` (on the browser thread): `PortalSession.take_typed_login()` hands the typed login over **once**, and only when the browser is past the login page (so a mistyped password is never offered). The question is then shown via `_on_ui_thread`, whatever happens next. **Bug fixed 2026-11-22:** the first version took the login inside `_scan` and only asked if a table was found. When the teacher clicked "The grade table is showing" while still logging in, the queued scan took the login, found no table, and the question never came (seen in the user's log). If the login differs from the saved one: "Save your portal password?" [Don't ask again] [Not now] [Save password], or "Update your saved password?" [Not now] [Update password]. Storage: `credentials.py` via `keyring` (pinned 25.7.0, `--hidden-import keyring.backends.Windows` in the build; the self-test writes/reads/deletes a dummy entry and the build fails without "password store: ok"). Corner menu (`_refresh_account`): username (or "Portal account" if only cookies are saved) → "Switch account" (confirm, warns to Save in the portal first; forgets the password, clears cookies + deletes `state.json`, reopens the portal if past Step 1) and "Forget my password". **Checked:** 5 browser tests against a fake portal + fake SSO (typed login noticed even while Python is idle; saved login signs in; wrong saved password pressed only once and gives up in < 8 s; switch account logs out); screens checked in a hidden Flet window. **Not yet tested on the real portal or Windows.** Unknown: does the real SSO keep the username after a wrong password (that's how a failure is detected quickly; otherwise it just waits 20 s)?

- **Update notice + self-update** (branch `update-notice`, 2026-11-21). At startup (packed Windows `.exe` only) a background thread asks `api.github.com/repos/Hala036/Grade-Filler/releases/latest`; if its tag is newer than `_version.__version__`, an amber bar shows "A new version of Grade Filler (X) is available. [Update] [Not now]". Update → confirm dialog (warns to click the portal's Save first, since the browser closes) → download to `GradeFiller.exe.new` with a progress bar → checks `MZ` header, size and the `sha256:` digest GitHub reports → renames the running `.exe` to `.old` (Windows allows renaming a running exe), moves the new one in, starts it with `PYINSTALLER_RESET_ENVIRONMENT=1`, closes via `_shut_down()`. On failure: "Try again" / "Download it myself"; the old `.exe` is put back if the swap fails. The next start deletes `.old`, retrying for ~30 s in its own thread (the old program may still be closing its browser). Offline at startup → nothing shown (logged). Folder not writable → Update opens the download in the browser instead. Version: the workflow writes `_version.py` from the tag; the self-test prints `version: X` and the build fails if it's wrong. HTTPS falls back to `certifi`'s certificates if the system's fail (python.org Python on macOS has none). Test builds are `0.0.0`, so they are always offered the latest release; handy for testing the swap. Run from source with `GRADEFILLER_CHECK_UPDATES=1` to see the bar (Update then just opens the download). **Not done (user agreed, 2026-11-21):** downloading into a hidden temp folder. `%TEMP%` can be on another drive than the `.exe`, which turns the instant rename into a slow copy that can fail halfway; the `.new` file next to the app for ~1 minute is the accepted cost. **Checked:** 14 unit tests; the live GitHub check (it returns v1.0.0, 117 MB, with digest); the bar's 3 states rendered with `ft.Screenshot`. **Not yet tested on Windows** (the real swap + restart).
  - **The teacher must download the first version that has this by hand, once.** v1.0.0 has no updater, so it can't announce anything.

- **App icon** (branch `windows-exe`): the user's A+ image, cropped and turned into `assets/icon.ico`; set on the `.exe` and the app window; the self-test fails if it isn't bundled. **Tested on the Windows computer (2026-11-20): everything works**; the icon showed in the taskbar but not on the desktop. Checked: the icon *is* embedded in the `.exe` (all 7 sizes extracted from the build). Cause: Windows' icon cache remembered the first, icon-less `GradeFiller.exe`. Fix: rename the file, or Win+R → `ie4uinit.exe -show`, or restart. A fresh laptop isn't affected.

- **GitHub build for the `.exe`** (branch `windows-exe`). `build-windows.yml` on `windows-latest`, Python 3.12: installs `requirements.txt` + `pyinstaller`, `flet pack main.py --name GradeFiller` (single file, no console), then runs `GradeFiller.exe --self-test` and fails unless Edge started from inside the bundle. Manual "Run workflow" (once on `main`) or any PR into `main` → download from the run's Artifacts; pushing tag `vX.Y.Z` also creates a Release, so the newest is always at `releases/latest/download/GradeFiller.exe`.

- **Uses the installed browser** (branch `windows-exe`). `PortalSession._launch_browser` tries `BROWSER_CHANNELS = ("msedge", "chrome", None)`: Edge (always on Windows), then Chrome (the user's Mac), then Playwright's downloaded Chromium. None found → "Couldn't open a web browser. Please install Microsoft Edge or Google Chrome and try again." Tested on the Mac (Edge missing → Chrome 157). Not yet tested with Edge on Windows. Each run is a clean browser profile; the login lives in the app's `state.json`.

- **Empty dropdowns look unfinished.** Every dropdown goes through `app._picker` / `_style_picker`: when nothing is chosen the label is faded (55% blue-grey), italic, with a light border; once chosen, label and border go dark and the value is semi-bold. Asked by the user because set and unset selectors looked alike.

- **Quieter column dropdowns.** Example values in brackets ("(now: 14)", "(e.g. 12)") now appear only for columns without a real name (portal "Column N", Excel "Unnamed: N"); named columns show just the name (user found it overwhelming once names were read). The Review step's 5-student preview is the sanity check. `matching.is_generic_label`.

- Deleted the obsolete `gui.py`, `upload_grades.py`, `save_session.py` (user agreed, 2026-11-19).

- "Show screen changes from the browser thread right away": **The screen froze on "Looking for the grade table…" until the window lost focus.** Cause: worker callbacks updated Flet from the Playwright thread; Flet's desktop transport queues outgoing changes in an `asyncio.Queue` (`flet_socket_server.send_message` → `put_nowait`), which doesn't wake the event loop when touched from another thread, so the new screen waited for any other event (focus change) to wake it. Fix: `PlaywrightWorker(dispatch=…)` hands every `on_success`/`on_error` to `GradeFillerApp._on_ui_thread`, which runs it via `page.run_task` on Flet's loop. Fill progress (`on_progress`, called on the browser thread) goes through the same path. **Rule for agents: never update Flet controls from the worker thread or any other thread; use `_on_ui_thread`.**

- **Round-up option in Match Columns: two joined arrows on a 0–1 line** (user's design, 2026-11-20). A "Round grades up" checkbox opens a number line (0, 0.1 … 1) with an amber arrow (start → joint) and a green arrow (joint → end); grades under an arrow go up to where it points. Default (0.1, 0.5, 1): .1–.4 → .5, .5–.9 → next whole. The three dots snap to tenths and can be dragged or clicked to; they push each other and never cross. Dragging the joint to 1 makes the green arrow disappear; dragging it back brings it back (no extra buttons, user's choice). Ends other than .5/1 are allowed (e.g. 0.3 → 0.6, 0.6 → 0.9) because the teacher "might" need them. A sentence underneath and on Review spells it out ("below 0.2 stays · 0.2–0.49 → 0.5 · 0.5–0.99 → 1"); examples use 0–1, not 12 (user's request). Each changed cell is listed under "Show details". Code: `gradefiller/round_line.py` (`RoundLine`, a Flet canvas + GestureDetector), `automation.RoundUp(start, joint, end)` in tenths + `apply_round_up`, applied in `build_fill_plan` (`FillPlan.rounded`). Only the **first** decimal digit counts (12.05 behaves like 12.0). Applied to the text Excel *displays*. Replaced two earlier versions (number boxes, then one clickable box per digit). Drawing checked with a Flet screenshot; dragging only checked in code, not by hand yet.

- "Choose the Mark Entry popup's table…": **The "Mark Entry" popup decides the table.** User fact: the portal's popup is always titled "Mark Entry" and holds only the table the teacher opened on purpose. Each discovered grid gets `titled=True` when it sits inside a *floating* box (CSS `position` absolute/fixed) whose text includes `POPUP_TITLE`, or is the only grid in a window titled that way, or sits in an iframe under such a box (`_frame_is_titled`). The floating requirement stops a page menu link saying "Mark Entry" from marking the table behind the popup. Exactly one titled table is chosen with no picker; otherwise the ID ranking below is the fallback. The ID count stays as a warning only: under 50% matched means "Only X of Y students are in your sheet. Is it the right sheet for this group?" (wrong section).
- "Pick the grade table automatically…": **Auto-select the right table + quieter screens.** `matching.rank_grids` scores every browser table by how many of its student IDs are in the chosen sheet (all IDs, not just the first row, since the portal and sheet sort differently), then by closeness of student count, then by "in front". The picker only opens when it isn't clear (no shared IDs, or a second table comes within 50%); otherwise Match Columns shows one line, "Filling the portal table with N students", with a "Wrong table?" link. The Student ID pairing collapses to "N of M students found" with "Change". Review shows one main line plus a one-line warning summary; ID lists, blanks and rows without an ID sit behind "Show details". Shorter texts throughout Match Columns.

- `7bc480c`: Copy grades as Excel shows them (they used to be rounded half-up to whole numbers).
- `91fcf04`: **Popup grade table.** The portal lists groups in one grid and opens the real grade table in a popup. The old scan mixed all grids on the page and read the group list. Now every grid (all tabs and iframes) is discovered and tagged with `data-gradefiller-grid`; the visible one in front wins; Match Columns has a "Table to fill" picker when there are several. The edit focus click moved from page centre to the row's ID cell.
- `85f61ab`: `requirements.txt` pins only direct deps (old `numpy==2.0.2` had no Python 3.13+ wheels, so the Windows install failed).
- `648351b`: Wizard rework. Flet 0.86 fixes (async `FilePicker.pick_files`, `Dropdown.on_select`, button `content`); one reused browser instead of two; retry/Back everywhere; worker survives callback errors; auto-matching; example values in dropdowns; validation; review step with preview; per-cell failures don't abort; end-of-run summary.

## Open issues / unverified (check on the REAL portal)

- [x] **"Looking for the grade table…" is slow on the real portal** (user, 2026-11-19). **Resolved:** the search itself was fast (the user confirmed with the timing log); the screen just didn't redraw until the window lost focus. See the thread fix in Recent fixes. Original notes: It's not the ID comparison: on a portal-sized fake page (110 students × 20 columns, 4 frames) the search takes ~30 ms, ranking ~2 ms and building the screen ~20 ms. Timing logs were added to find out (since then they go only to the file, not the terminal, at the user's request; size-capped at ~1 MB): `GradeFiller/gradefiller.log` (in `%APPDATA%` on Windows, `~/Library/Application Support` on macOS) show the wait for the browser worker, ms per tab/frame (with URL; `[Mark Entry]` marks the popup), `save_session` time, and how the table was chosen ("by the 'Mark Entry' popup title" vs "by student IDs"). Next: read a real log. Suspects: a slow or still-loading frame (`frame.evaluate` has no timeout), slow `innerText` on a big portal page in `gfTitled`, or the worker still busy with an earlier browser action.

- [x] **Popup detection + Mark Entry auto-pick: confirmed on the real portal** (user, 2026-11-19). The Mark Entry table is chosen without asking.
- [x] **Decimals:** the portal accepts values like `36.8` (user, 2026-11-19).
- [x] **Edit focus click on the ID cell: confirmed on the real portal** (user, 2026-11-20). Grades are filled correctly and stay in their cells; each value is committed with Tab.
- [x] **Header labels: confirmed on the real portal** (user, 2026-11-20). The popup's two-row header (`Mark` spans First 25 / Second 25 / Third 10) is laid out honouring rowspan/colspan in `_DISCOVER_JS`; each column takes its lowest name.
- [x] Merged `improving-UI` → `main` after a successful real run (2026-11-20).
- [ ] Only the updater has automated tests (`tests/`). Earlier checks were throwaway scripts: fake DHTMLX pages + headless Chromium, and a stub Flet page driving each step. Worth adding under `tests/`.
- [x] Fixed: the `#` row-number column (values 1, 2, 3 …) and unnamed empty spacer columns are no longer offered as fill targets; name matching uses the max mark as a tie-breaker (`First 25` → `First/25`, not `First/30`) (`matching._split_label`, `_similarity`, `portal_fillable_columns`).

## Next features

0. **v1.1.0 is released but not yet tested on Windows. Don't send the teacher the link until it is.** The teacher (on v1.0.0, which has no updater) downloads v1.1.0 by hand once: https://github.com/Hala036/Grade-Filler/releases/latest/download/GradeFiller.exe
1. **Test auto-login on Windows** with v1.1.0 (or a 0.0.0 test build; then do this *before* clicking Update): first run → Open the portal → log in → "Save your password?" should appear in the app at once → Save. Close, reopen → the portal should log in by itself. Try "Switch account" and "Forget my password". Check Windows Credential Manager shows "Grade Filler (portal.alquds.edu)".
2. **Test the update notice:** run a 0.0.0 test build (any PR/manual build's Artifacts) on the Windows computer. The bar offers 1.1.0; Update should download, swap and reopen as v1.1.0, and the `.old` file should vanish within ~30 s.
3. Whatever the user asks for next. Add it here.
   - Possible: remember the round-up setting between runs (today it resets each launch).
   - The local sample sheet `2251-1214311-sec1-Grades.xlsx` (gitignored) was removed from the folder on 2026-11-20; tests since then use made-up data.
