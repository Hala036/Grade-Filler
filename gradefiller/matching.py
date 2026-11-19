"""Suggest column mappings automatically. Pure functions, no UI or browser.

Nothing is tied to a course or exam: the Student ID pair is found by
comparing actual values (which portal column and which Excel column share
the most student IDs), and grade pairs are suggested by name similarity.
Suggestions are only a starting point; the user can change any of them.
"""
from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
from typing import Optional

import pandas as pd

from gradefiller.automation import GridColumn, looks_numeric, normalize_id

# IDs are long digit strings; this keeps small grade values (e.g. "10")
# from being mistaken for matching IDs.
_MIN_ID_LENGTH = 5
_NAME_MATCH_THRESHOLD = 0.75


def _id_set(values) -> set[str]:
    return {v for v in (normalize_id(x) for x in values) if len(v) >= _MIN_ID_LENGTH}


def detect_id_columns(
    grid: list[GridColumn], df: pd.DataFrame
) -> tuple[Optional[int], Optional[str], int]:
    """Return (portal_index, excel_column, shared_id_count) with most overlap."""
    best = (None, None, 0)
    excel_sets = {c: _id_set(df[c]) for c in df.columns}
    for g in grid:
        portal = _id_set(g.values)
        if not portal:
            continue
        for col, ex in excel_sets.items():
            n = len(portal & ex)
            if n > best[2]:
                best = (g.index, col, n)
    return best


def id_overlap(grid_column: GridColumn, df: pd.DataFrame, excel_column: str) -> tuple[int, int]:
    """(portal IDs found in the Excel column, total portal IDs)."""
    portal = {v for v in (normalize_id(x) for x in grid_column.values) if v}
    excel = {v for v in (normalize_id(x) for x in df[excel_column]) if v}
    return len(portal & excel), len(portal)


def grade_candidate_columns(df: pd.DataFrame, id_column: Optional[str]) -> list[str]:
    """Excel columns that could hold grades: numeric, and not a copy of the ID column.

    Teachers' sheets often repeat the name/ID block next to each exam
    (e.g. 'ID number.1'); those copies are hidden from the grade choices.
    """
    ids = _id_set(df[id_column]) if id_column in df.columns else set()
    out = []
    for c in df.columns:
        if c == id_column or not looks_numeric(df[c].tolist()):
            continue
        if ids:
            col_ids = _id_set(df[c])
            if col_ids and len(col_ids & ids) >= 0.5 * len(col_ids):
                continue
        out.append(c)
    return out


def portal_fillable_columns(grid: list[GridColumn], id_index: Optional[int]) -> list[GridColumn]:
    """Portal columns that could receive grades (not the ID, not text like names)."""
    out = []
    for g in grid:
        if g.index == id_index:
            continue
        non_blank = [v for v in g.values if v]
        if non_blank and not looks_numeric(non_blank):
            continue  # names, statuses, etc.
        out.append(g)
    return out


def _norm_label(s: str) -> str:
    s = unicodedata.normalize("NFKC", str(s)).lower()
    s = re.sub(r"\.\d+$", "", s)       # pandas duplicate suffix: "Quiz.1"
    s = re.sub(r"[/(]\s*\d+\s*\)?", " ", s)  # max-mark suffix: "First/25", "(25)"
    s = re.sub(r"[^\w]+", " ", s)
    return " ".join(s.split())


def _similarity(a: str, b: str) -> float:
    a, b = _norm_label(a), _norm_label(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b.split() or b in a.split():
        return 0.9
    return SequenceMatcher(None, a, b).ratio()


def suggest_grade_mappings(
    portal_columns: list[GridColumn], excel_columns: list[str]
) -> list[tuple[int, str]]:
    """One-to-one (portal_index, excel_column) pairs whose names clearly match."""
    scored = []
    for g in portal_columns:
        if g.label.startswith("Column ") and g.label[7:].isdigit():
            continue  # generic label, nothing to compare
        for c in excel_columns:
            s = _similarity(g.label, c)
            if s >= _NAME_MATCH_THRESHOLD:
                scored.append((s, g.index, c))

    scored.sort(key=lambda t: -t[0])
    used_p, used_e, pairs = set(), set(), []
    for _, p, c in scored:
        if p in used_p or c in used_e:
            continue
        used_p.add(p)
        used_e.add(c)
        pairs.append((p, c))
    pairs.sort(key=lambda t: t[0])  # left-to-right portal order
    return pairs
