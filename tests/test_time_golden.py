# SPDX-License-Identifier: Apache-2.0
"""The calendar's answers, frozen.

Every question the engine asks its calendar — the labels of a window,
the buckets a coarser grain groups it into, the day-count weights, the
fields of ``mo.time``, the short first period, the header's Excel
spelling, the totals interleave — is asked here over a matrix of
windows and compared with the answers recorded in
``golden/time_calendar.json``. The file is the gate under which the
calendar's code may move: a refactor that changes any answer for the
grains covered here is a behaviour change and must say so by
regenerating the file on purpose::

    MODELEON_REGEN_TIME_GOLDEN=1 pytest tests/test_time_golden.py
"""
import json
import os
from datetime import date
from pathlib import Path
from typing import Any, Dict, List

import pytest

from modeleon.compile.excel import header as H
from modeleon.compile.excel.totals import build_totals_plan, drawable_kinds, header_rows
from modeleon.core import time as T

GOLDEN = Path(__file__).parent / "golden" / "time_calendar.json"
REGEN = os.environ.get("MODELEON_REGEN_TIME_GOLDEN") == "1"

#: (grain, start, n, first_day) — whole windows, windows crossing a
#: year end, leap days, one-period windows, and short first periods.
WINDOWS = [
    ("day", "2026-01-05", 40, None),
    ("day", "2025-12-25", 62, None),
    ("day", "2024-02-27", 5, None),
    ("day", "2026-03-30", 4, None),
    ("day", "2026-01-01", 1, None),
    ("month", "2026-01", 13, None),
    ("month", "2025-11", 5, None),
    ("month", "2024-12", 1, None),
    ("month", "2026-01", 26, None),
    ("month", "2026-02", 3, None),
    ("quarter", "2026-Q1", 8, None),
    ("quarter", "2025-Q4", 3, None),
    ("quarter", "2026-Q3", 5, None),
    ("quarter", "2026-Q1", 5, date(2026, 2, 1)),
    ("quarter", "2026-Q2", 4, date(2026, 5, 1)),
    ("quarter", "2025-Q4", 2, date(2025, 12, 1)),
    ("year", "2026", 3, None),
    ("year", "2024", 2, None),
    ("year", "2026", 3, date(2026, 5, 1)),
    ("year", "2026", 2, date(2026, 12, 1)),
    ("year", "2025", 1, date(2025, 2, 1)),
    ("week", "2026-W01", 10, None),
    ("week", "2026-W51", 6, None),
    ("week", "2025-W52", 4, None),
    ("iso_year", "2026", 3, None),
    # Hours: a whole day from midnight, across a day end, a month end, a
    # year end and a leap day, and one hour alone.
    ("hour", "2026-10-21T00", 24, None),
    ("hour", "2026-10-21T22", 30, None),
    ("hour", "2026-01-31T22", 5, None),
    ("hour", "2025-12-31T22", 5, None),
    ("hour", "2024-02-28T23", 26, None),
    ("hour", "2026-03-29T01", 3, None),
    ("hour", "2026-10-21T13", 1, None),
]

STARTS = [
    date(2026, 1, 1), date(2026, 2, 1), date(2026, 4, 1), date(2026, 5, 17),
    date(2026, 7, 1), date(2026, 12, 31), date(2024, 2, 29),
]

KINDS = [("quarter", "year"), ("year", "quarter"), ("quarter",), ("year",),
         ("month", "quarter"), ("month",), ("day", "month"),
         ("week",), ("iso_year",), ("week", "iso_year"), ("week", "month"), ("year", "iso_year"),
         ("day",), ("day", "week"), ("month", "day", "year")]


def _iso(v: Any) -> Any:
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (list, tuple)):
        return [_iso(x) for x in v]
    if isinstance(v, dict):
        return {str(k): _iso(x) for k, x in v.items()}
    return v


def _ask(fn, *args, **kwargs) -> Any:
    try:
        return _iso(fn(*args, **kwargs))
    except (ValueError, KeyError, TypeError) as e:
        return {"error": f"{type(e).__name__}: {e}"}


def _cell_chain(opening: bool):
    """A fake anchor sheet: row letters S/E/D/M/I at column j+3; an
    opening cell when asked."""
    def cell(field_: str, j: int):
        if j < 0:
            return "O9" if opening else None
        return f"{field_[0].upper()}{j + 3}"
    return cell


def _window_answers(grain: str, start: str, n: int, first_day: Any) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    out["labels"] = {s: _ask(T._period_labels, start, None, n, grain, s) for s in T.TIME_LABEL_STYLES}
    out["weights"] = _ask(T.period_weights, grain, start, n, first_day)
    out["fields"] = {f: _ask(T.time_field_series, f, start, grain, n, first_day) for f in T.TIME_FIELDS}
    anchor = T._parse(start, grain)
    out["first_day_of_start"] = _iso(T._first_day(anchor, grain))
    out["containing"] = _ask(T._containing_label, first_day or T._first_day(anchor, grain), grain)
    out["days_in_first"] = _ask(T._days_in_period, anchor, grain)
    out["step"] = _ask(lambda: T._format(T._step(anchor, grain), grain, "iso"))
    out["count_to_last"] = _ask(lambda: T._count(anchor, T._parse(out["labels"]["iso"][-1], grain), grain))
    window = T.DefaultWindow(start, grain, n, None, first_day)
    out["spelled_first_day"] = _ask(T.spelled_first_day, window)
    out["spelled_first_day_sourced"] = _ask(T.spelled_first_day, T.DefaultWindow(start, grain, n, True, first_day))
    coarser = [g for g in T.GRAINS if T._ORDER[g] > T._ORDER[grain]]
    out["targets"] = {}
    for target in coarser:
        out["targets"][target] = {
            "buckets": _ask(T.bucket_ranges, grain, target, start, n),
            "coarse_first_day": _ask(T.coarse_first_day, start, grain, first_day, target),
            "target_label_of_start": _ask(T._target_label, anchor, grain, target),
            "regrain_sum": _ask(T.regrain_series, [float(i + 1) for i in range(n)], grain, target, start, "sum", first_day),
            "regrain_mean": _ask(T.regrain_series, [float(i % 3) for i in range(n)], grain, target, start, "mean", first_day),
        }
    out["spell_inline"] = {
        f"{f}@{i}": _ask(H.spell_inline, f, i, grain, "B2", first_day)
        for f in T.TIME_FIELDS for i in (0, 1, 2)
    }
    out["spell_anchor"] = {
        f"{f}@{i}{'+o' if opening else ''}": _ask(H.spell_anchor, f, i, grain, "B2", _cell_chain(opening), first_day)
        for f in T.TIME_FIELDS for i in (0, 1, 2) for opening in (False, True)
    }
    out["totals"] = {}
    out["totals_notes"] = {}
    labels = out["labels"]["finance"]
    for kinds in KINDS:
        # As the layout draws them: only the kinds this grain nests in,
        # nesting with each other; the rest are notes.
        kept, notes = drawable_kinds(kinds, grain)
        if notes:
            out["totals_notes"]["+".join(kinds)] = notes
        for stride in ((1, 3) if kinds == ("quarter", "year") else (1,)):
            plan = build_totals_plan(kept, grain, start, n, 3, stride)
            key = "+".join(kinds) + f"/s{stride}"
            if plan is None:
                out["totals"][key] = None
                continue
            out["totals"][key] = {
                "kinds": list(plan.kinds),
                "columns": [list(c) for c in plan.columns],
                "col_of": [[list(k), v] for k, v in plan.col_of.items()],
                "buckets": _iso(plan.buckets),
                "depth": plan.depth,
                # The tiered header is long for a wide window; the
                # columns above already pin its geometry there.
                "header_rows": header_rows(plan, labels) if isinstance(labels, list) and n <= 13 else None,
            }
    return out


def _answers() -> Dict[str, Any]:
    golden: Dict[str, Any] = {"grains": list(T.GRAINS), "order": dict(T._ORDER),
                              "months_in_period": dict(T.MONTHS_IN_PERIOD),
                              "step_months": dict(H.STEP), "windows": {}, "starts": {}}
    for grain, start, n, first_day in WINDOWS:
        key = f"{grain}|{start}|{n}|{first_day.isoformat() if first_day else '-'}"
        golden["windows"][key] = _window_answers(grain, start, n, first_day)
    for grain in T.GRAINS:
        golden["starts"][grain] = {
            d.isoformat(): {
                "window_start_label": _ask(T.window_start_label, d, grain),
                "stub_first_day": _ask(T.stub_first_day, d, grain),
                "containing_label": _ask(T._containing_label, d, grain),
            }
            for d in STARTS
        }
        golden["starts"][grain]["labels"] = {
            lab: _ask(T._parse, lab, grain)
            for lab in ("2026-01-05", "2026-01", "2026-Q1", "2026", "2026-W01", "2026-H1",
                        "2026-01-05T13", "garbage")
        }
    return golden


def test_the_calendars_answers_are_the_recorded_ones():
    answers = _answers()
    if REGEN or not GOLDEN.exists():
        GOLDEN.write_text(json.dumps(answers, indent=0, sort_keys=True, default=str) + "\n")
        if REGEN:
            pytest.skip("golden regenerated")
    recorded = json.loads(GOLDEN.read_text())
    # Compare per window so a mismatch names the window and the question.
    assert answers["grains"] == recorded["grains"]
    assert answers["order"] == recorded["order"]
    assert answers["months_in_period"] == recorded["months_in_period"]
    assert answers["step_months"] == recorded["step_months"]
    for key, got in answers["windows"].items():
        want = recorded["windows"][key]
        for question in want:
            assert json.loads(json.dumps(got[question], default=str)) == want[question], (key, question)
    assert json.loads(json.dumps(answers["starts"], default=str)) == recorded["starts"]
    assert set(answers["windows"]) == set(recorded["windows"])
