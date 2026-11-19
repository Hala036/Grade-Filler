"""A dedicated worker thread that owns the Playwright session end-to-end.

Playwright's sync API pins itself to the OS thread that first called
sync_playwright().start() — every subsequent call (page navigation,
clicking, closing) must happen on that same thread, or Playwright raises
"cannot switch to a different thread". Since the wizard's UI needs to kick
off portal actions from multiple places (button clicks, dialog confirms,
which Flet may dispatch on different threads), this worker exists as the
single place all Playwright calls run: everything is submitted as a task
and executed serially on one thread for the lifetime of the browser.

The reverse direction matters too: the callbacks that update the screen
must NOT run on this thread. Flet queues screen changes in an asyncio
queue that only wakes its event loop when touched from the loop's own
thread; from here, a new screen sat unseen until some other event (e.g.
the window losing focus) woke the loop. So results are handed back
through `dispatch`, which runs callbacks on Flet's event loop.
"""
from __future__ import annotations

import queue
import threading
import traceback
from typing import Any, Callable, Optional

Dispatch = Callable[..., None]
"""dispatch(callback, *args): run callback(*args) on the UI thread."""


def _call_here(callback: Callable, *args: Any) -> None:
    callback(*args)


class PlaywrightWorker:
    def __init__(self, dispatch: Optional[Dispatch] = None):
        self._dispatch: Dispatch = dispatch or _call_here
        self._tasks: "queue.Queue[tuple[Callable, Callable, Callable]]" = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self):
        while True:
            fn, on_success, on_error = self._tasks.get()
            if fn is None:
                return
            # A failing task or callback must never kill this thread: every
            # later browser action would then silently queue up and never run.
            try:
                try:
                    result = fn()
                except Exception as e:
                    if on_error:
                        self._dispatch(on_error, e)
                else:
                    if on_success:
                        self._dispatch(on_success, result)
            except Exception:
                traceback.print_exc()

    def submit(self, fn: Callable, on_success: Callable = None, on_error: Callable = None):
        """Queue fn to run on the worker thread. Non-blocking.

        on_success(result) / on_error(exception) are handed to `dispatch`
        (the UI thread in the app) once fn finishes, so they may update the
        screen but must not call Playwright directly; submit another task
        for that.
        """
        self._tasks.put((fn, on_success, on_error))

    def stop(self):
        self._tasks.put((None, None, None))
