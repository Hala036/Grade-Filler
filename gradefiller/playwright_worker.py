"""A dedicated worker thread that owns the Playwright session end-to-end.

Playwright's sync API pins itself to the OS thread that first called
sync_playwright().start() — every subsequent call (page navigation,
clicking, closing) must happen on that same thread, or Playwright raises
"cannot switch to a different thread". Since the wizard's UI needs to kick
off portal actions from multiple places (button clicks, dialog confirms,
which Flet may dispatch on different threads), this worker exists as the
single place all Playwright calls run: everything is submitted as a task
and executed serially on one thread for the lifetime of the browser.
"""
from __future__ import annotations

import queue
import threading
import traceback
from typing import Callable


class PlaywrightWorker:
    def __init__(self):
        self._tasks: "queue.Queue[tuple[Callable, Callable, Callable]]" = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self):
        while True:
            fn, on_success, on_error = self._tasks.get()
            if fn is None:
                return
            # A failing callback must never kill this thread: every later
            # browser action would then silently queue up and never run.
            try:
                try:
                    result = fn()
                except Exception as e:
                    if on_error:
                        on_error(e)
                else:
                    if on_success:
                        on_success(result)
            except Exception:
                traceback.print_exc()

    def submit(self, fn: Callable, on_success: Callable = None, on_error: Callable = None):
        """Queue fn to run on the worker thread. Non-blocking.

        on_success(result) / on_error(exception) are called from the same
        worker thread once fn finishes — if they touch Flet controls, that
        is fine under Flet 0.80+'s threading model (direct control updates
        from a background thread are supported), but they must NOT call
        back into Playwright objects owned by a *different* worker.
        """
        self._tasks.put((fn, on_success, on_error))

    def stop(self):
        self._tasks.put((None, None, None))
