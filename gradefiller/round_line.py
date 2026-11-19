"""The round-up number line: two joined arrows on a 0-1 scale.

The amber arrow starts at `start` and points at `joint`; the green one
starts at `joint` and points at `end`. Everything under an arrow is rounded
up to where it points. All three points can be dragged (or clicked to);
they snap to tenths and push each other so they never cross.
"""
from __future__ import annotations

from typing import Callable, Optional

import flet as ft
import flet.canvas as cv

from gradefiller.automation import RoundUp

AMBER, AMBER_DARK = ft.Colors.AMBER_400, ft.Colors.AMBER_800
GREEN, GREEN_DARK = ft.Colors.GREEN_400, ft.Colors.GREEN_700
AXIS, LABEL = ft.Colors.BLUE_GREY_200, ft.Colors.BLUE_GREY_600

WIDTH, HEIGHT = 580, 76
PAD_X = 22  # room for the end labels and handles
LINE_Y = 30
HANDLE_R = 9
HEAD_LEN, HEAD_HALF = 15, 10
SHAFT = 8


def tenth_label(n: int) -> str:
    return "0" if n == 0 else "1" if n == 10 else f"0.{n}"


def _paint(color, width: float = 0, fill: bool = False) -> ft.Paint:
    return ft.Paint(
        color=color,
        stroke_width=width,
        stroke_cap=ft.StrokeCap.ROUND,
        style=ft.PaintingStyle.FILL if fill else ft.PaintingStyle.STROKE,
    )


class RoundLine:
    def __init__(self, rule: RoundUp, on_change: Callable[[RoundUp], None]):
        self.points = [rule.start, rule.joint, rule.end]
        self.on_change = on_change
        self._dragging: Optional[int] = None  # handle index being dragged
        self._pending: list[int] = []  # coincident handles; picked by drag direction
        self.canvas = cv.Canvas(shapes=self._shapes(), width=WIDTH, height=HEIGHT)
        self.control = ft.GestureDetector(
            content=self.canvas,
            mouse_cursor=ft.MouseCursor.CLICK,
            on_pan_start=self._pan_start,
            on_pan_update=self._pan_update,
            on_pan_end=self._pan_end,
            on_tap_up=self._tap,
            drag_interval=20,
        )

    @property
    def rule(self) -> RoundUp:
        return RoundUp(*self.points)

    # ---------- geometry ----------

    @staticmethod
    def x(n: float) -> float:
        return PAD_X + n * (WIDTH - 2 * PAD_X) / 10

    @staticmethod
    def tenth_at(x: float) -> int:
        n = round((x - PAD_X) / ((WIDTH - 2 * PAD_X) / 10))
        return max(0, min(10, n))

    # ---------- moving points ----------

    def _set(self, i: int, n: int) -> None:
        """Move point i to n, pushing the others so the order holds."""
        p = self.points
        if p[i] == n:
            return
        p[i] = n
        for j in range(i + 1, 3):
            p[j] = max(p[j], n)
        for j in range(i - 1, -1, -1):
            p[j] = min(p[j], n)
        self.canvas.shapes = self._shapes()
        self.canvas.update()
        self.on_change(self.rule)

    def _nearest(self, x: float) -> list[int]:
        """Indexes of the handle(s) closest to x (several when they overlap)."""
        dist = [abs(self.x(n) - x) for n in self.points]
        best = min(dist)
        return [i for i, d in enumerate(dist) if d - best < 0.5]

    @staticmethod
    def _pick(candidates: list[int], towards_right: bool) -> int:
        # Stacked handles: pulling right takes the last one, left the first,
        # so dragging always opens up the arrows instead of getting stuck.
        return candidates[-1] if towards_right else candidates[0]

    def _pan_start(self, e: ft.DragStartEvent):
        cands = self._nearest(e.local_position.x)
        if len(cands) == 1:
            self._dragging, self._pending = cands[0], []
        else:
            self._dragging, self._pending = None, cands

    def _pan_update(self, e: ft.DragUpdateEvent):
        n = self.tenth_at(e.local_position.x)
        if self._dragging is None:
            if not self._pending:
                return
            here = self.points[self._pending[0]]
            if n == here:
                return
            self._dragging = self._pick(self._pending, n > here)
            self._pending = []
        self._set(self._dragging, n)

    def _pan_end(self, _e):
        self._dragging, self._pending = None, []

    def _tap(self, e: ft.TapEvent):
        n = self.tenth_at(e.local_position.x)
        cands = self._nearest(e.local_position.x)
        here = self.points[cands[0]]
        if n != here:
            self._set(self._pick(cands, n > here), n)

    # ---------- drawing ----------

    def _arrow(self, a: int, b: int, color, dark) -> list[cv.Shape]:
        if a >= b:
            return []
        x1, tip = self.x(a), self.x(b) - HANDLE_R  # stop where the next handle starts
        base = tip - HEAD_LEN
        shapes = []
        if base > x1:
            shapes.append(cv.Line(x1, LINE_Y, base + 1, LINE_Y, paint=_paint(color, SHAFT)))
        shapes.append(
            cv.Path(
                [
                    cv.Path.MoveTo(tip, LINE_Y),
                    cv.Path.LineTo(base, LINE_Y - HEAD_HALF),
                    cv.Path.LineTo(base, LINE_Y + HEAD_HALF),
                    cv.Path.Close(),
                ],
                paint=_paint(dark, fill=True),
            )
        )
        return shapes

    def _handle(self, n: int, color) -> list[cv.Shape]:
        return [
            cv.Circle(self.x(n), LINE_Y, HANDLE_R, paint=_paint(ft.Colors.WHITE, fill=True)),
            cv.Circle(self.x(n), LINE_Y, HANDLE_R, paint=_paint(color, 3.5)),
        ]

    def _shapes(self) -> list[cv.Shape]:
        start, joint, end = self.points
        amber_on, green_on = start < joint, joint < end
        shapes: list[cv.Shape] = [cv.Line(self.x(0), LINE_Y, self.x(10), LINE_Y, paint=_paint(AXIS, 2))]
        for n in range(11):
            key = n in self.points and (amber_on or green_on)
            shapes.append(cv.Line(self.x(n), LINE_Y - 6, self.x(n), LINE_Y + 6, paint=_paint(AXIS, 2)))
            shapes.append(
                cv.Text(
                    self.x(n),
                    LINE_Y + 16,
                    tenth_label(n),
                    style=ft.TextStyle(
                        size=12,
                        color=ft.Colors.BLUE_GREY_900 if key else LABEL,
                        weight=ft.FontWeight.BOLD if key else ft.FontWeight.NORMAL,
                    ),
                    alignment=ft.Alignment.TOP_CENTER,
                )
            )
        shapes += self._arrow(start, joint, AMBER, AMBER_DARK)
        shapes += self._arrow(joint, end, GREEN, GREEN_DARK)

        # Handles. The joint is the amber arrow's end and the green one's start;
        # the end handle is only drawn while the green arrow is there (dragging
        # the joint back brings it back).
        shapes += self._handle(start, AMBER_DARK)
        if green_on:
            shapes += self._handle(end, GREEN_DARK)
        shapes += self._handle(joint, GREEN_DARK if green_on else AMBER_DARK)
        return shapes
