I'm continuing work on Grade Filler, a Python desktop app (Flet UI + Playwright) that types grades from an Excel sheet into a university's grades portal (portal.alquds.edu, a DHTMLX grid). Repo: https://github.com/Hala036/Grade-Filler, branch `improving-UI` (the newest version; `main` is old).

Before doing anything, read `CLAUDE.local.md` in the repo root. It has the file layout, the decisions I've made, recent fixes, open issues and the feature plan. Keep that file updated as you work: when you fix something or find a new issue, record it there before you finish. It's committed in the repo (branch `improving-UI`); if you can't access the repo, ask me to paste it.

Key points in case you can't see the file:
- The app is used on a **Windows laptop by a non-technical person**: no git, no terminal if at all possible. Keep UI text plain and friendly.
- It must work for **any course and any number of exams**. Nothing hardcoded; columns are read from the live grid and matched in the UI.
- Grades are copied **exactly as Excel displays them** (no rounding). Blank cells are left unchanged by default. Duplicate IDs and non-numeric cells are skipped and reported.
- In the portal, the real grade table opens in a **popup** over a list of groups; the app detects separate grids and picks the one in front.
- Still unverified on the real portal: popup detection, the edit focus click (now on the row's ID cell), whether the portal accepts decimals like 36.8, and header-label reading.
- The old `GUI.py`, `upload_grades.py` and `save_session.py` are obsolete.

Top priority next: **a Windows .exe that needs no terminal.** Plan: a GitHub Actions `windows-latest` build with `flet pack` from `main.py`, downloadable from GitHub in the browser, launching the laptop's own Edge via `channel="msedge"` so no Chromium download is needed. Ask me before pushing anything.

What I want to work on now: <describe the feature or bug here>
