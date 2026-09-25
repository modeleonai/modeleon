# SPDX-License-Identifier: Apache-2.0
"""Arithmetic and comparisons on date values follow Excel.

In Excel a date IS a number: the count of days since 1899-12-30, with
the time of day as a fraction (2024-03-31 is 45382, noon that day is
45382.5). The written workbook computes on those numbers, so the value
Python computes for the same formula has to be the same number, the
same date, or the same TRUE/FALSE.
"""

from datetime import date, datetime, timedelta

import pytest

import modeleon as mo

D0 = date(2024, 1, 15)   # serial 45306
D1 = date(2024, 3, 31)   # serial 45382
NOON = datetime(2024, 3, 31, 12)   # serial 45382.5


def V(value):
    return mo.Variable(value)


class TestShiftByDays:
    def test_date_plus_whole_number_is_a_date(self):
        r = V(D0) + 5
        assert r.value == date(2024, 1, 20)
        assert type(r.value) is date
        assert r.value_type == "datetime"

    def test_number_plus_date_is_a_date(self):
        assert (5 + V(D0)).value == date(2024, 1, 20)
        assert (V(5) + V(D0)).value == date(2024, 1, 20)

    def test_date_minus_whole_number_is_a_date(self):
        assert (V(D0) - 15).value == date(2023, 12, 31)
        assert (V(D0) - V(15)).value == date(2023, 12, 31)

    def test_whole_float_keeps_a_pure_date(self):
        r = V(D0) + 5.0
        assert r.value == date(2024, 1, 20)
        assert type(r.value) is date

    def test_fractional_days_give_a_datetime(self):
        assert (V(D0) + 0.5).value == datetime(2024, 1, 15, 12)
        assert (V(D0) - 0.25).value == datetime(2024, 1, 14, 18)
        assert (V(D0) + 0.5).value_type == "datetime"

    def test_datetime_keeps_its_time_of_day(self):
        r = V(datetime(2024, 1, 15, 9, 30)) + 1
        assert r.value == datetime(2024, 1, 16, 9, 30)

    def test_timedelta_shifts_by_its_days(self):
        assert (V(D0) + timedelta(days=5)).value == date(2024, 1, 20)
        assert (timedelta(days=5) + V(D0)).value == date(2024, 1, 20)
        assert (V(D0) - timedelta(days=15)).value == date(2023, 12, 31)
        assert (V(D0) + V(timedelta(days=5))).value == date(2024, 1, 20)

    def test_timedelta_with_hours_gives_a_datetime(self):
        assert (V(D0) + timedelta(hours=12)).value == datetime(2024, 1, 15, 12)
        assert (V(NOON) + timedelta(hours=6)).value == datetime(2024, 3, 31, 18)

    def test_timedelta_without_a_date_counts_as_days(self):
        assert (V(5) + timedelta(days=1)).value == 6
        assert (V(1) + timedelta(hours=12)).value == 1.5

    def test_list_forms(self):
        assert (V([D0, D1]) + 1).value == [date(2024, 1, 16), date(2024, 4, 1)]
        assert (V([D0, D1]) + V([1, 0.5])).value == [
            date(2024, 1, 16), datetime(2024, 3, 31, 12),
        ]
        assert (V([D0, D1]) + timedelta(days=1)).value == [
            date(2024, 1, 16), date(2024, 4, 1),
        ]

    def test_holes_and_errors_pass_through(self):
        assert (V([D0, None]) + 1).value == [date(2024, 1, 16), None]
        assert (V([D0, "#N/A"]) + 1).value == [date(2024, 1, 16), "#N/A"]

    def test_text_is_not_a_day_count(self):
        assert (V(D0) + "x").value == "#VALUE!"


class TestDaysBetween:
    def test_date_minus_date_is_whole_days(self):
        r = V(D1) - V(D0)
        assert r.value == 76
        assert type(r.value) is int
        assert r.value_type != "datetime"

    def test_datetime_minus_date_keeps_the_time_of_day(self):
        r = V(NOON) - V(date(2024, 1, 1))
        assert r.value == 90.5
        assert r.value_type != "datetime"

    def test_a_datetime_gives_float_days(self):
        r = V(datetime(2024, 3, 31)) - date(2024, 1, 1)
        assert r.value == 90.0
        assert type(r.value) is float

    def test_list_form(self):
        assert (V([D1, D1]) - V([D0, D1])).value == [76, 0]

    def test_number_minus_date_is_a_number(self):
        r = 45400 - V(D1)
        assert r.value == 18
        assert r.value_type != "datetime"


class TestOtherArithmeticIsOnTheSerial:
    def test_multiply(self):
        r = V(D1) * 2
        assert r.value == 90764
        assert r.value_type != "datetime"
        assert (2 * V(D1)).value == 90764

    def test_divide(self):
        r = V(D1) / 7
        assert r.value == pytest.approx(45382 / 7)
        assert r.value_type != "datetime"
        assert (90764 / V(D1)).value == 2.0

    def test_power_and_mod(self):
        assert (V(D1) ** 2).value == 45382 ** 2
        assert (V(D1) % 7).value == 45382 % 7

    def test_datetime_serial_carries_the_fraction(self):
        assert (V(NOON) * 2).value == 90765.0

    def test_date_plus_date_is_a_number(self):
        r = V(D0) + V(D1)
        assert r.value == 45306 + 45382
        assert r.value_type != "datetime"

    def test_divide_by_zero(self):
        assert (V(D1) / 0).value == "#DIV/0!"

    def test_divide_by_a_zero_serial(self):
        # 1899-12-30 is serial 0 and an empty timedelta is 0 days.
        assert (5 / V(date(1899, 12, 30))).value == "#DIV/0!"
        assert (V(5) / timedelta(0)).value == "#DIV/0!"

    def test_list_form(self):
        r = V([D0, D1]) * 2
        assert r.value == [90612, 90764]
        assert r.value_type != "datetime"


class TestComparisons:
    @pytest.mark.parametrize("op, expected", [
        ("<", True), ("<=", True), (">", False),
        (">=", False), ("==", False), ("!=", True),
    ])
    def test_date_vs_date(self, op, expected):
        r = _compare(V(D0), op, V(D1))
        assert r.value is expected

    def test_equal_dates(self):
        assert (V(D0) == V(D0)).value is True
        assert (V(D0) != V(D0)).value is False
        assert (V(D0) <= V(D0)).value is True

    def test_date_vs_serial_number(self):
        assert (V(D1) > 45000).value is True
        assert (V(D1) < 45000).value is False
        assert (V(D1) == 45382).value is True
        assert (V(D1) != 45382).value is False
        assert (V(D1) >= 45382).value is True
        assert (V(D1) <= 45381).value is False
        assert (V(D1) == 5).value is False
        assert (V(D1) == V(45382)).value is True

    def test_number_on_the_left(self):
        assert (45000 < V(D1)).value is True
        assert (45382 == V(D1)).value is True

    def test_datetime_vs_number_uses_the_time_of_day(self):
        assert (V(NOON) > 45382).value is True
        assert (V(NOON) == 45382.5).value is True
        assert (V(NOON) > V(D1)).value is True

    def test_date_vs_timedelta_days(self):
        assert (V(D1) > timedelta(days=45000)).value is True

    @pytest.mark.parametrize("op, expected", [
        ("==", False), ("!=", True), ("<", True),
        ("<=", True), (">", False), (">=", False),
    ])
    def test_date_vs_text_numbers_sort_first(self, op, expected):
        r = _compare(V(D1), op, "x")
        assert r.value is expected

    def test_date_vs_logical_sorts_after_text(self):
        assert (V(D1) < True).value is True
        assert (V(D1) == True).value is False  # noqa: E712 - the operator is the subject

    def test_list_forms(self):
        assert (V([D0, D1]) < V([D1, D0])).value == [True, False]
        assert (V([D0, D1]) == 45306).value == [True, False]
        assert (V([D0, D1]) > D0).value == [False, True]

    def test_holes_and_errors_pass_through(self):
        assert (V([D0, None]) < D1).value == [True, None]
        assert (V([D0, "#N/A"]) < D1).value == [True, "#N/A"]


class TestIfOverADateComparison:
    def test_false_comparison_takes_the_else_branch(self):
        assert mo.IF(V(D1) == 5, 1, 0).value == 0

    def test_true_comparison_takes_the_then_branch(self):
        assert mo.IF(V(D1) > V(D0), 1, 0).value == 1
        assert mo.IF(V(D1) >= 45382, "on or after", "before").value == "on or after"

    def test_element_wise(self):
        assert mo.IF(V([D0, D1]) > 45350, 1, 0).value == [0, 1]


def _compare(left, op, right):
    return {
        "<": lambda: left < right,
        "<=": lambda: left <= right,
        ">": lambda: left > right,
        ">=": lambda: left >= right,
        "==": lambda: left == right,
        "!=": lambda: left != right,
    }[op]()
