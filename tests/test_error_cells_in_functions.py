# SPDX-License-Identifier: Apache-2.0
"""An error cell and a hole reach an aggregate and a math function as the
spreadsheet reads them.

Arithmetic carries an error token through (``'#DIV/0!'``) and a hole (a
month not entered yet) as None; ``SUM``, ``MAX``, ``MIN``, ``AVERAGE`` and
``ABS`` raised Python's TypeError on either — a division by zero upstream
failed at a SUM three blocks away, and so did the SUM of every plan/actual
row with an actual month still open. What the spreadsheet does was measured
in Excel 16.95: the first error in reading order is the result, a blank is
skipped, nothing left is 0 (AVERAGE: ``#DIV/0!``).
"""

from __future__ import annotations

import pytest

import modeleon as mo

AGGREGATES = (mo.SUM, mo.MAX, mo.MIN, mo.AVERAGE)


def _with_an_error() -> tuple[mo.Model, mo.Variable]:
    model = mo.Model("M", default_grain="year", default_start="2026", default_periods=3)
    with model:
        model.a = mo.Variable([1.0, 0.0, 2.0], var_type="list")
        model.q = 1.0 / model.a
    return model, model.q


def test_an_aggregate_over_an_error_cell_is_that_error() -> None:
    _model, q = _with_an_error()
    assert q.value == [1.0, "#DIV/0!", 0.5]
    for fn in AGGREGATES:
        assert fn(q).value == "#DIV/0!", fn


@pytest.mark.parametrize("fn", AGGREGATES)
def test_the_first_error_in_reading_order_wins(fn) -> None:
    assert fn(mo.Variable(["#N/A", "#DIV/0!", 3.0])).value == "#N/A"
    assert fn(mo.Variable(["#DIV/0!", "#N/A", 3.0])).value == "#DIV/0!"
    assert fn(mo.Variable([1.0, 2.0]), mo.Variable(["#VALUE!", "#N/A"])).value == "#VALUE!"


def test_a_hole_is_skipped_as_a_blank_is() -> None:
    row = mo.Variable([1.0, None, 2.0])
    assert mo.SUM(row).value == 3.0
    assert mo.MAX(row).value == 2.0
    assert mo.MIN(row).value == 1.0
    assert mo.AVERAGE(row).value == 1.5


def test_nothing_left_is_zero_and_an_average_of_nothing_divides_by_zero() -> None:
    row = mo.Variable([None, None, None])
    assert mo.SUM(row).value == 0
    assert mo.MAX(row).value == 0
    assert mo.MIN(row).value == 0
    assert mo.AVERAGE(row).value == "#DIV/0!"


def test_an_actual_month_not_entered_yet_does_not_break_a_total() -> None:
    model = mo.Model(
        "M", default_grain="month", default_start="2026-01", default_periods=3,
        tracks=mo.Tracks("plan", "actual"),
    )
    with model:
        model.r = mo.Variable(plan=[10.0, 20.0, 30.0], actual=[11.0, None, None])
        model.z = mo.Variable(plan=[10.0, 20.0, 30.0], actual=[None, None, None])
    assert mo.SUM(model.r).value["actual"] == 11.0
    assert mo.SUM(model.r).value["plan"] == 60.0
    assert mo.AVERAGE(model.r).value["actual"] == 11.0
    assert mo.SUM(model.z).value["actual"] == 0
    assert mo.AVERAGE(model.z).value["actual"] == "#DIV/0!"


def test_abs_of_an_error_cell_is_that_error() -> None:
    _model, q = _with_an_error()
    assert mo.ABS(q).value == [1.0, "#DIV/0!", 0.5]
    assert mo.ROUND(q, 1).value == [1.0, "#DIV/0!", 0.5]


def test_an_aggregate_of_numbers_is_what_it_was() -> None:
    model = mo.Model("M", default_grain="year", default_start="2026", default_periods=3)
    with model:
        model.a = mo.Variable([1.0, -4.0, 2.0], var_type="list")
    assert mo.SUM(model.a).value == -1.0
    assert mo.MAX(model.a).value == 2.0 and mo.MIN(model.a).value == -4.0
    assert mo.AVERAGE(model.a).value == -1.0 / 3.0
    assert mo.ABS(model.a).value == [1.0, 4.0, 2.0]


# ─── An error is never TRUE or FALSE ──────────────────────────────
# Measured in Excel 16.95: IF(1/0 > 0.2, 5, 0), AND(FALSE, 1/0),
# OR(TRUE, 1/0), NOT(1/0) and CHOOSE(1/0, 1, 2) are all #DIV/0!; an
# error in the branch IF does not take, or the choice CHOOSE does not
# pick, changes nothing; ISBLANK(1/0) is FALSE. Before, an error
# condition read as TRUE: a bonus over an average with a zero month
# showed 50,000 where the workbook shows #DIV/0!.


def test_a_condition_over_an_aggregate_error_is_that_error() -> None:
    _model, q = _with_an_error()
    assert mo.IF(mo.AVERAGE(q) > 0.2, 50_000, 0).value == "#DIV/0!"
    assert mo.AND(mo.SUM(q) > 1, True).value == "#DIV/0!"


def test_a_condition_row_carries_its_error_period_by_period() -> None:
    _model, q = _with_an_error()
    assert mo.IF(q > 0.7, 1, 0).value == [1, "#DIV/0!", 0]
    assert mo.NOT(q > 0.7).value == [False, "#DIV/0!", True]


def test_the_first_error_decides_and_and_or_whatever_the_others_say() -> None:
    err = mo.Variable("#DIV/0!")
    assert mo.AND(False, err).value == "#DIV/0!"
    assert mo.OR(True, err).value == "#DIV/0!"
    assert mo.OR(mo.Variable("#N/A"), err).value == "#N/A"


def test_an_error_in_the_branch_not_taken_changes_nothing() -> None:
    err = mo.Variable("#DIV/0!")
    assert mo.IF(False, err, 5).value == 5
    assert mo.IF(True, err, 5).value == "#DIV/0!"
    assert mo.CHOOSE(1, 5, err).value == 5
    assert mo.CHOOSE(2, 5, err).value == "#DIV/0!"


def test_choose_by_an_error_is_that_error_not_value() -> None:
    assert mo.CHOOSE(mo.Variable("#DIV/0!"), 1, 2).value == "#DIV/0!"


def test_isblank_of_an_error_is_false() -> None:
    assert mo.ISBLANK(mo.Variable("#DIV/0!")).value is False


def test_not_of_an_error_is_that_error() -> None:
    assert mo.NOT(mo.Variable("#DIV/0!")).value == "#DIV/0!"


def test_a_tracked_condition_carries_its_error_per_track() -> None:
    model = mo.Model(
        "M", default_grain="month", default_start="2026-01", default_periods=3,
        tracks=mo.Tracks("plan", "actual"),
    )
    with model:
        model.rev = mo.Variable(plan=[100.0, 100.0, 100.0], actual=[90.0, 0.0, None])
        model.cost = mo.Variable(plan=[60.0, 60.0, 60.0], actual=[50.0, 40.0, None])
        model.bonus = mo.IF(model.cost / model.rev < 0.6, 1, 0)
    assert model.bonus.value["plan"] == [0, 0, 0]
    assert model.bonus.value["actual"][:2] == [1, "#DIV/0!"]
