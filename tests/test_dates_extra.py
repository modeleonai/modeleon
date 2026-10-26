# SPDX-License-Identifier: Apache-2.0
"""``DATEDIF``, ``WEEKDAY``, ``WEEKNUM``, ``NETWORKDAYS`` and ``WORKDAY``
compute what the spreadsheet computes, quirks included, with the
formula as the product.

``DATEDIF`` keeps Excel's own arithmetic for the ``"MD"`` and ``"YD"``
units — the ones Microsoft's help flags — so the workbook recalculates
to the engine's numbers: ``"MD"`` moves the start's day into the month
before the end's and lets it roll over as ``DATE()`` does (31 January
to 1 March is -1), ``"YD"`` moves the end's month and day into the
start's year, not the start's into the end's (3 January 2015 to 28 July
2016 is 206, not 207). Each function is one call per period through the
shared helper: a scalar date stands in every period, a series gives a
series, tracks lift, an error cell passes through; a blank or a value
that is not a date is refused by name at build time, as the other date
helpers refuse it. The holidays of ``NETWORKDAYS`` / ``WORKDAY`` are one
range every period reads, never a series.
"""

from __future__ import annotations

from datetime import date, datetime

import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall, Literal, VarRef
from modeleon.functions.dates import DATEDIF, NETWORKDAYS, WEEKDAY, WEEKNUM, WORKDAY
from modeleon.functions.registry import (
    admits_in_loop,
    excel_name,
    file_prefix,
    is_elementwise,
    is_range_taking,
    range_positions,
    spec,
)

# Microsoft's own DATEDIF examples.
S, E = date(2001, 6, 1), date(2002, 8, 15)


def V(value, **kwargs):
    return mo.Variable(value, **kwargs)


# ══════════════════════════════════════════════════════════════════════
# DATEDIF
# ══════════════════════════════════════════════════════════════════════


class TestDatedif:
    @pytest.mark.parametrize("unit, expected", [
        ("Y", 1), ("M", 14), ("D", 440), ("MD", 14), ("YM", 2), ("YD", 75),
    ])
    def test_the_documented_examples(self, unit, expected):
        assert DATEDIF(S, E, unit).value == expected
        assert DATEDIF(V(S), V(E), unit).value == expected

    def test_whole_years(self):
        assert DATEDIF(date(2001, 1, 1), date(2003, 1, 1), "Y").value == 2
        assert DATEDIF(date(2001, 1, 2), date(2003, 1, 1), "Y").value == 1
        # 29 February's anniversary is 1 March: a year is not complete on the 28th.
        assert DATEDIF(date(2024, 2, 29), date(2025, 2, 28), "Y").value == 0
        assert DATEDIF(date(2024, 2, 29), date(2025, 3, 1), "Y").value == 1

    def test_whole_months_need_the_day_reached(self):
        # 31 January to 28 February is no whole month to the spreadsheet.
        assert DATEDIF(date(2025, 1, 31), date(2025, 2, 28), "M").value == 0
        assert DATEDIF(date(2025, 3, 31), date(2025, 4, 30), "M").value == 0
        assert DATEDIF(date(2025, 1, 31), date(2025, 4, 30), "M").value == 2
        assert DATEDIF(date(2025, 1, 15), date(2025, 2, 15), "M").value == 1

    def test_months_past_the_years(self):
        assert DATEDIF(date(2001, 6, 15), date(2002, 6, 10), "YM").value == 11
        assert DATEDIF(date(2001, 6, 15), date(2002, 6, 15), "YM").value == 0
        assert DATEDIF(date(2001, 1, 15), date(2001, 12, 20), "YM").value == 11

    @pytest.mark.parametrize("start, end, expected", [
        # The known quirk: the start's day lands in a February that is
        # too short and rolls over into March, past the end.
        (date(2016, 1, 31), date(2016, 3, 1), -1),
        (date(2025, 1, 31), date(2025, 3, 1), -2),
        (date(2025, 1, 29), date(2025, 3, 1), 0),
        (date(2012, 1, 28), date(2012, 3, 1), 2),
        (date(2011, 1, 29), date(2011, 3, 1), 0),
        # The month before January is the December before.
        (date(2023, 12, 31), date(2024, 1, 15), 15),
        (date(2023, 11, 30), date(2024, 1, 15), 16),
        # The end's day reached: a plain difference.
        (date(2024, 1, 15), date(2024, 5, 20), 5),
        (date(2024, 1, 15), date(2024, 5, 15), 0),
    ])
    def test_days_past_the_months_quirks_included(self, start, end, expected):
        assert DATEDIF(start, end, "MD").value == expected

    @pytest.mark.parametrize("start, end, expected", [
        # The end's month and day in the start's year: February 2015 has
        # 28 days, so 206 where a count near the end date gives 207.
        (date(2015, 1, 3), date(2016, 7, 28), 206),
        # The same rule across 23 years: 323, where the count near the
        # end date crosses 29 February 2012 and gives 324.
        (date(1988, 6, 22), date(2012, 5, 11), 323),
        # The end's day before the start's in the year: the year after.
        (date(2010, 10, 11), date(2012, 3, 1), 141),
        (date(2010, 10, 11), date(2011, 3, 1), 141),
        # A whole number of years is 0 days, leap years included.
        (date(2003, 3, 1), date(2004, 3, 1), 0),
        (date(2004, 1, 1), date(2005, 1, 1), 0),
        (date(2004, 3, 1), date(2005, 3, 1), 0),
        (date(2020, 2, 29), date(2021, 2, 28), 365),
        # 29 February placed in a common year rolls over to 1 March.
        (date(2015, 1, 3), date(2016, 2, 29), 57),
        # A 29 February at either end lies outside the values checked in Excel:
        # the rule applied as written (LibreOffice gives 0 and 2 here).
        (date(2024, 2, 29), date(2025, 3, 1), 1),
        (date(2015, 2, 28), date(2016, 3, 1), 1),
    ])
    def test_days_past_the_years_quirks_included(self, start, end, expected):
        assert DATEDIF(start, end, "YD").value == expected

    @pytest.mark.parametrize("unit", ["Y", "M", "D", "MD", "YM", "YD"])
    def test_equal_dates_are_zero(self, unit):
        assert DATEDIF(S, S, unit).value == 0

    def test_the_unit_is_read_in_any_case_and_written_as_given(self):
        assert DATEDIF(S, E, "md").value == 14
        assert DATEDIF(S, E, "yd").value == 75
        start, end = V(S, display_name="start"), V(E, display_name="end")
        assert DATEDIF(start, end, "md").formula == "DATEDIF(start, end, 'md')"
        assert DATEDIF(start, end, V("YM", display_name="unit")).value == 2

    def test_a_start_after_the_end_is_the_period_s_num(self):
        assert DATEDIF(E, S, "D").value == "#NUM!"
        ends = V([date(2001, 5, 1), date(2001, 7, 1), date(2002, 8, 15)])
        assert DATEDIF(S, ends, "M").value == ["#NUM!", 1, 14]
        guarded = mo.IF(ends >= S, DATEDIF(S, ends, "M"), 0)
        assert guarded.value == [0, 1, 14]

    @pytest.mark.parametrize("unit", ["X", "DM", "", 3, None])
    def test_a_unit_the_spreadsheet_lacks_is_refused(self, unit):
        with pytest.raises((ValueError, TypeError), match="DATEDIF"):
            DATEDIF(S, E, unit)

    def test_element_wise_over_series(self):
        ends = V([date(2024, 1, 31), date(2024, 2, 29), date(2024, 3, 31)])
        starts = V([date(2024, 1, 1), date(2024, 1, 1), date(2024, 2, 1)])
        start = V(date(2024, 1, 1))
        assert DATEDIF(start, ends, "D").value == [30, 59, 90]
        assert DATEDIF(starts, ends, "D").value == [30, 59, 59]
        assert DATEDIF(date(2024, 1, 1), ends, "D").var_type == 'list'
        assert DATEDIF(V([date(2024, 1, 1)]), ends, "D").value == [30, 59, 90]
        assert DATEDIF(starts, ends, ["D", "M", "MD"]).value == [30, 1, 30]

    def test_series_of_different_lengths_are_refused_as_the_operators_refuse_them(self):
        with pytest.raises(ValueError, match="different lengths"):
            DATEDIF(V([S, S, S]), V([E, E]), "D")

    def test_an_error_cell_is_the_period_s_result(self):
        ends = V([E, "#N/A", "#DIV/0!"])
        assert DATEDIF(S, ends, "D").value == [440, "#N/A", "#DIV/0!"]
        assert DATEDIF(V("#N/A"), E, "D").value == "#N/A"

    def test_a_blank_is_refused_by_name_with_its_period(self):
        with pytest.raises(TypeError, match="DATEDIF needs a date: 'leave' is blank"):
            DATEDIF(S, V(None, display_name="leave"), "M")
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=2)
        with m:
            m.leave = mo.Variable([date(2026, 1, 31), None], display_name="leave")
        with pytest.raises(TypeError, match="'leave' is blank in Feb 2026"):
            DATEDIF(date(2026, 1, 1), m.leave, "D")

    def test_a_value_that_is_not_a_date_is_refused_at_build_time(self):
        with pytest.raises(TypeError, match="Expected a date"):
            DATEDIF(45000, E, "D")
        with pytest.raises(ValueError):
            DATEDIF("not a date", E, "D")

    def test_a_datetime_counts_by_its_date_and_an_iso_string_is_a_date(self):
        assert DATEDIF(datetime(2001, 6, 1, 23, 30), datetime(2002, 8, 15, 1), "D").value == 440
        assert DATEDIF("2001-06-01", "2002-08-15", "YD").value == 75

    def test_tracked_operands_lift_per_track(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"),
                     default_grain="month", default_start="2026-01", default_periods=2)
        with m:
            m.paid = mo.Variable(plan=[date(2026, 1, 31), date(2026, 2, 28)],
                                 actual=[date(2026, 2, 3), date(2026, 3, 2)])
            m.late = DATEDIF(date(2026, 1, 1), m.paid, "D")
        assert dict(m.late._value.items()) == {"plan": [30, 58], "actual": [33, 60]}
        assert m.late.var_type == 'list'

    def test_a_whole_number_without_a_unit_and_the_ast_is_a_call(self):
        m = mo.Model("m")
        start = m.start = V(S)
        end = m.end = V(E, unit="$")
        out = DATEDIF(start, end, "M")
        assert out.value_type == 'int' and out.var_type == 'scalar'
        assert out._unit is None
        assert out.formula == "DATEDIF(start, end, 'M')"
        assert out.source_code == "DATEDIF(start, end, 'M')"
        expr = out._expr
        assert isinstance(expr, FuncCall) and expr.func == "DATEDIF"
        assert isinstance(expr.args[0], VarRef) and expr.args[0].var is start
        assert isinstance(expr.args[1], VarRef) and expr.args[1].var is end
        assert isinstance(expr.args[2], Literal) and expr.args[2].value == "M"
        assert expr.render_backends == frozenset({'excel'})
        assert out.dependencies == {start.id, end.id}

    def test_an_axised_operand_is_refused(self):
        region = mo.Variable(["N", "S"], display_name="region")
        by_region = mo.Variable([S, E], indexed_by=[region])
        with pytest.raises(ValueError, match="DATEDIF"):
            DATEDIF(by_region, E, "D")


# ══════════════════════════════════════════════════════════════════════
# WEEKDAY and WEEKNUM
# ══════════════════════════════════════════════════════════════════════

# Monday 1 January 2024 to Sunday 7 January 2024.
WEEK = [date(2024, 1, d) for d in range(1, 8)]


class TestWeekday:
    @pytest.mark.parametrize("return_type, expected", [
        (1, [2, 3, 4, 5, 6, 7, 1]),
        (2, [1, 2, 3, 4, 5, 6, 7]),
        (3, [0, 1, 2, 3, 4, 5, 6]),
        (11, [1, 2, 3, 4, 5, 6, 7]),
        (12, [7, 1, 2, 3, 4, 5, 6]),
        (13, [6, 7, 1, 2, 3, 4, 5]),
        (14, [5, 6, 7, 1, 2, 3, 4]),
        (15, [4, 5, 6, 7, 1, 2, 3]),
        (16, [3, 4, 5, 6, 7, 1, 2]),
        (17, [2, 3, 4, 5, 6, 7, 1]),
    ])
    def test_every_return_type(self, return_type, expected):
        assert WEEKDAY(V(WEEK), return_type).value == expected
        assert [WEEKDAY(d, return_type).value for d in WEEK] == expected

    def test_sunday_is_one_when_no_type_is_given(self):
        assert WEEKDAY(date(2024, 1, 7)).value == 1
        assert WEEKDAY(V(WEEK)).value == [2, 3, 4, 5, 6, 7, 1]
        assert WEEKDAY(V(WEEK), 1.9).value == [2, 3, 4, 5, 6, 7, 1]

    def test_the_type_is_written_only_when_given(self):
        d = V(date(2024, 1, 7), display_name="d")
        assert WEEKDAY(d).formula == "WEEKDAY(d)"
        assert WEEKDAY(d, 2).formula == "WEEKDAY(d, 2)"
        assert len(WEEKDAY(d)._expr.args) == 1

    @pytest.mark.parametrize("return_type", [0, 4, 10, 18, 21, -1])
    def test_a_type_the_spreadsheet_lacks_is_num(self, return_type):
        with pytest.raises(ValueError, match="WEEKDAY.*#NUM!"):
            WEEKDAY(date(2024, 1, 7), return_type)

    def test_a_type_that_is_not_a_number_is_value(self):
        with pytest.raises(TypeError, match="WEEKDAY.*#VALUE!"):
            WEEKDAY(date(2024, 1, 7), "two")
        with pytest.raises(TypeError, match="WEEKDAY.*#VALUE!"):
            WEEKDAY(date(2024, 1, 7), True)

    def test_a_datetime_counts_by_its_date(self):
        assert WEEKDAY(datetime(2024, 1, 7, 23, 59)).value == 1

    def test_a_blank_and_an_error_cell(self):
        with pytest.raises(TypeError, match="WEEKDAY needs a date"):
            WEEKDAY(V(None))
        assert WEEKDAY(V([date(2024, 1, 7), "#N/A"])).value == [1, "#N/A"]

    def test_a_whole_number_without_a_unit(self):
        out = WEEKDAY(V(WEEK, unit="$"))
        assert out.value_type == 'int' and out.var_type == 'list' and out._unit is None


class TestWeeknum:
    def test_week_one_holds_the_first_of_january(self):
        # 1 January 2024 is a Monday: Sunday the 7th begins week 2 of a
        # Sunday-started count and ends week 1 of a Monday-started one.
        assert WEEKNUM(date(2024, 1, 1)).value == 1
        assert WEEKNUM(date(2024, 1, 6)).value == 1
        assert WEEKNUM(date(2024, 1, 7)).value == 2
        assert WEEKNUM(date(2024, 1, 7), 1).value == 2
        assert WEEKNUM(date(2024, 1, 7), 2).value == 1
        assert WEEKNUM(date(2024, 1, 8), 2).value == 2
        # 1 January 2021 is a Friday: Sunday the 3rd begins week 2.
        assert WEEKNUM(date(2021, 1, 1)).value == 1
        assert WEEKNUM(date(2021, 1, 2)).value == 1
        assert WEEKNUM(date(2021, 1, 3)).value == 2
        assert WEEKNUM(date(2021, 1, 3), 2).value == 1
        assert WEEKNUM(date(2021, 1, 4), 2).value == 2

    def test_the_last_days_of_a_year_reach_week_53_or_54(self):
        # 365 days are 52 weeks and a day: 31 December is always past
        # week 52, and a leap year that opens on the last day of a week
        # reaches 54 (1 January 2000 was a Saturday).
        assert WEEKNUM(date(2024, 12, 29)).value == 53
        assert WEEKNUM(date(2024, 12, 31)).value == 53
        assert WEEKNUM(date(2024, 12, 31), 2).value == 53
        assert WEEKNUM(date(2023, 12, 31)).value == 53
        assert WEEKNUM(date(2023, 12, 31), 2).value == 53
        assert WEEKNUM(date(2023, 12, 24), 2).value == 52
        assert WEEKNUM(date(2000, 12, 31)).value == 54
        assert WEEKNUM(date(2000, 12, 31), 2).value == 53

    def test_the_last_days_of_a_year_in_every_mode(self):
        # 2021 opened and closed on a Friday: whichever day a week begins
        # on, 31 December lies past 52 whole weeks from the week of 1
        # January. LibreOffice numbers these days 1 (see test_excel_dates).
        for return_type in (1, 2, 11, 12, 13, 14, 15, 16, 17):
            assert WEEKNUM(date(2021, 12, 31), return_type).value == 53, return_type
        assert WEEKNUM(date(2021, 12, 31), 21).value == 52
        # 2028 is a leap year opening on a Saturday: 54 for a Sunday-started
        # count, 53 when the week begins on Saturday or Monday.
        assert WEEKNUM(date(2028, 12, 31), 1).value == 54
        assert WEEKNUM(date(2028, 12, 31), 17).value == 54
        assert WEEKNUM(date(2028, 12, 31), 16).value == 53
        assert WEEKNUM(date(2028, 12, 31), 2).value == 53
        assert WEEKNUM(date(2028, 12, 31), 21).value == 52
        # A Thursday-started week 53 of 2024 opens on Thursday 26 December.
        assert WEEKNUM(date(2024, 12, 25), 14).value == 52
        assert WEEKNUM(date(2024, 12, 26), 14).value == 53

    @pytest.mark.parametrize("return_type, expected", [
        (11, [1, 1, 1, 1, 1, 1, 1]),
        (12, [1, 2, 2, 2, 2, 2, 2]),
        (13, [1, 1, 2, 2, 2, 2, 2]),
        (14, [1, 1, 1, 2, 2, 2, 2]),
        (15, [1, 1, 1, 1, 2, 2, 2]),
        (16, [1, 1, 1, 1, 1, 2, 2]),
        (17, [1, 1, 1, 1, 1, 1, 2]),
    ])
    def test_a_week_that_begins_on_any_day(self, return_type, expected):
        assert WEEKNUM(V(WEEK), return_type).value == expected

    def test_type_21_is_the_iso_week(self):
        assert WEEKNUM(date(2021, 1, 1), 21).value == 53
        assert WEEKNUM(date(2021, 1, 4), 21).value == 1
        assert WEEKNUM(date(2024, 12, 30), 21).value == 1
        assert WEEKNUM(date(2024, 12, 29), 21).value == 52
        assert WEEKNUM(V(WEEK), 21).value == [1, 1, 1, 1, 1, 1, 1]

    def test_the_type_is_written_only_when_given(self):
        d = V(date(2024, 1, 7), display_name="d")
        assert WEEKNUM(d).formula == "WEEKNUM(d)"
        assert WEEKNUM(d, 21).formula == "WEEKNUM(d, 21)"

    @pytest.mark.parametrize("return_type", [0, 3, 10, 18, 20, 22, 150])
    def test_a_type_the_spreadsheet_lacks_is_num(self, return_type):
        with pytest.raises(ValueError, match="WEEKNUM.*#NUM!"):
            WEEKNUM(date(2024, 1, 7), return_type)

    def test_a_blank_an_error_cell_and_the_result_s_type(self):
        with pytest.raises(TypeError, match="WEEKNUM needs a date"):
            WEEKNUM(V([date(2024, 1, 7), None]))
        out = WEEKNUM(V([date(2024, 1, 7), "#VALUE!"], unit="$"))
        assert out.value == [2, "#VALUE!"]
        assert out.value_type == 'int' and out._unit is None


# ══════════════════════════════════════════════════════════════════════
# NETWORKDAYS and WORKDAY
# ══════════════════════════════════════════════════════════════════════

# Microsoft's own examples.
NW_START, NW_END = date(2012, 10, 1), date(2013, 3, 1)
NW_HOLIDAYS = [date(2012, 11, 22), date(2012, 12, 4), date(2013, 1, 21)]
WD_START, WD_DAYS = date(2008, 10, 1), 151
WD_HOLIDAYS = [date(2008, 11, 26), date(2008, 12, 4), date(2009, 1, 21)]


class TestNetworkdays:
    def test_the_documented_examples(self):
        assert NETWORKDAYS(NW_START, NW_END).value == 110
        assert NETWORKDAYS(NW_START, NW_END, NW_HOLIDAYS[0]).value == 109
        assert NETWORKDAYS(NW_START, NW_END, NW_HOLIDAYS).value == 107
        assert NETWORKDAYS(V(NW_START), V(NW_END), V(NW_HOLIDAYS)).value == 107

    def test_both_ends_count_and_weekends_do_not(self):
        jan = (date(2024, 1, 1), date(2024, 1, 31))      # Monday to Wednesday
        assert NETWORKDAYS(*jan).value == 23
        assert NETWORKDAYS(date(2024, 1, 1), date(2024, 1, 1)).value == 1
        assert NETWORKDAYS(date(2024, 1, 6), date(2024, 1, 6)).value == 0
        assert NETWORKDAYS(date(2024, 1, 6), date(2024, 1, 7)).value == 0
        assert NETWORKDAYS(date(2024, 1, 5), date(2024, 1, 8)).value == 2

    def test_a_start_after_the_end_counts_negative(self):
        assert NETWORKDAYS(date(2024, 1, 31), date(2024, 1, 1)).value == -23
        assert NETWORKDAYS(NW_END, NW_START, NW_HOLIDAYS).value == -107

    def test_holidays_count_once_on_weekdays_inside_the_span(self):
        jan = (date(2024, 1, 1), date(2024, 1, 31))
        assert NETWORKDAYS(*jan, [date(2024, 1, 1), date(2024, 1, 1)]).value == 22
        assert NETWORKDAYS(*jan, [date(2024, 1, 6), date(2024, 1, 7)]).value == 23
        assert NETWORKDAYS(*jan, [date(2023, 12, 29), date(2024, 2, 1)]).value == 23
        assert NETWORKDAYS(*jan, [date(2024, 1, 1), date(2024, 1, 31)]).value == 21
        assert NETWORKDAYS(*jan, ()).value == 23
        assert NETWORKDAYS(*jan, []).value == 23

    def test_holidays_may_be_blank_dated_text_or_a_line_of_dates(self):
        jan = (date(2024, 1, 1), date(2024, 1, 31))
        assert NETWORKDAYS(*jan, [date(2024, 1, 1), None]).value == 22
        assert NETWORKDAYS(*jan, ["2024-01-01", datetime(2024, 1, 2, 9)]).value == 21
        assert NETWORKDAYS(*jan, V([date(2024, 1, 1), None, date(2024, 1, 2)])).value == 21
        assert NETWORKDAYS(*jan, V(date(2024, 1, 1))).value == 22

    def test_a_zero_among_the_holidays_is_skipped_as_the_spreadsheet_skips_it(self):
        jan = (date(2024, 1, 1), date(2024, 1, 31))
        assert NETWORKDAYS(*jan, [date(2024, 1, 1), 0, 0.0]).value == 22
        assert WORKDAY(jan[0], 1, [0, date(2024, 1, 2)]).value == date(2024, 1, 3)
        # Holidays in a longer window: a line extended with mo.zero()
        # fills the rest of the window with 0.0, which Excel reads as a
        # serial before every date.
        m = mo.Model("m", default_grain="month", default_start="2024-01", default_periods=7)
        with m:
            m.hol = mo.Variable([date(2024, 1, 1), date(2024, 1, 2)], extend=mo.zero())
        assert m.hol.value[2:] == [0.0] * 5
        assert NETWORKDAYS(*jan, m.hol).value == 21

    def test_a_shorter_holidays_line_keeps_its_own_length_with_extend_none(self):
        m = mo.Model("m", default_grain="month", default_start="2024-01", default_periods=7)
        with m:
            m.hol = mo.Variable([date(2024, 1, 1), date(2024, 1, 2)], extend=mo.none())
            m.n = NETWORKDAYS(date(2024, 1, 1), date(2024, 1, 31), m.hol)
        assert m.hol.value == [date(2024, 1, 1), date(2024, 1, 2)]
        assert m.n.value == 21
        assert isinstance(m.n._expr.args[2], VarRef) and m.n.dependencies == {m.hol.id}

    def test_a_non_date_among_the_holidays_is_refused_by_name(self):
        jan = (date(2024, 1, 1), date(2024, 1, 31))
        m = mo.Model("m", default_grain="month", default_start="2024-01", default_periods=3)
        with m:
            m.hol = mo.Variable([date(2024, 1, 1), 45000, None], display_name="hol")
        with pytest.raises(TypeError, match="NETWORKDAYS holidays must be dates: 'hol' holds "
                                            "45000 in Feb 2024"):
            NETWORKDAYS(*jan, m.hol)
        with pytest.raises(TypeError, match="the list holds 'x' at position 2"):
            NETWORKDAYS(*jan, [date(2024, 1, 1), "x"])
        with pytest.raises(TypeError, match="WORKDAY holidays must be dates: the holidays holds "
                                            "45000"):
            WORKDAY(jan[0], 1, 45000)
        with pytest.raises(TypeError, match="holds True"):
            NETWORKDAYS(*jan, [True])

    def test_a_list_of_lines_is_refused_with_a_pointer_to_one_line(self):
        a, b = V(date(2024, 1, 1)), V(date(2024, 1, 2))
        with pytest.raises(TypeError, match="a list of lines is not a range"):
            NETWORKDAYS(date(2024, 1, 1), date(2024, 1, 31), [a, b])
        with pytest.raises(TypeError, match="a list of lines is not a range"):
            WORKDAY(date(2024, 1, 1), 1, (a,))

    def test_an_error_among_the_holidays_is_the_result(self):
        jan = (date(2024, 1, 1), date(2024, 1, 31))
        assert NETWORKDAYS(*jan, V([date(2024, 1, 1), "#N/A"])).value == "#N/A"
        assert NETWORKDAYS(V([jan[0], jan[0]]), jan[1], ["#REF!"]).value == ["#REF!", "#REF!"]

    def test_element_wise_over_the_dates_with_one_set_of_holidays(self):
        ends = V([date(2024, 1, 31), date(2024, 2, 29), date(2024, 3, 29)])
        hol = V([date(2024, 1, 1), date(2024, 2, 5), date(2024, 3, 15)])
        out = NETWORKDAYS(date(2024, 1, 1), ends, hol)
        assert out.value == [22, 42, 62]
        assert out.var_type == 'list'
        starts = V([date(2024, 1, 1), date(2024, 2, 1), date(2024, 3, 1)])
        assert NETWORKDAYS(starts, ends, hol).value == [22, 20, 20]
        assert NETWORKDAYS(ends, date(2024, 1, 1)).value == [-23, -44, -65]
        assert NETWORKDAYS(date(2024, 1, 1), V([date(2024, 1, 31), "#N/A"])).value == [23, "#N/A"]

    def test_blanks_and_non_dates_in_the_span_are_refused_by_name(self):
        with pytest.raises(TypeError, match="NETWORKDAYS needs a date: 'end' is blank"):
            NETWORKDAYS(date(2024, 1, 1), V(None, display_name="end"))
        with pytest.raises(TypeError, match="Expected a date"):
            NETWORKDAYS(date(2024, 1, 1), 45000)

    def test_tracked_dates_lift_per_track(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"),
                     default_grain="month", default_start="2024-01", default_periods=2)
        with m:
            m.paid = mo.Variable(plan=[date(2024, 1, 31), date(2024, 2, 29)],
                                 actual=[date(2024, 2, 2), date(2024, 3, 1)])
            m.hol = mo.Variable([date(2024, 1, 1)])
            m.days = NETWORKDAYS(date(2024, 1, 1), m.paid, m.hol)
        assert dict(m.days._value.items()) == {"plan": [22, 43], "actual": [24, 44]}

    def test_the_holidays_are_one_operand_of_the_call_not_a_series(self):
        m = mo.Model("m")
        start = m.start = V(date(2024, 1, 1))
        ends = m.ends = V([date(2024, 1, 31), date(2024, 2, 29)], unit="$")
        hol = m.hol = V([date(2024, 1, 1), date(2024, 2, 5)])
        out = NETWORKDAYS(start, ends, hol)
        assert out.formula == "NETWORKDAYS(start, ends, hol)"
        assert out.source_code == "NETWORKDAYS(start, ends, hol)"
        expr = out._expr
        assert isinstance(expr, FuncCall) and expr.func == "NETWORKDAYS"
        assert [type(a) for a in expr.args] == [VarRef, VarRef, VarRef]
        assert expr.args[2].var is hol
        assert expr.render_backends == frozenset({'excel'})
        assert out.dependencies == {start.id, ends.id, hol.id}
        assert out.value_type == 'int' and out._unit is None
        bare = NETWORKDAYS(start, ends)
        assert bare.formula == "NETWORKDAYS(start, ends)" and len(bare._expr.args) == 2
        literal = NETWORKDAYS(start, ends, [date(2024, 1, 1)])
        assert isinstance(literal._expr.args[2], Literal)
        assert literal._expr.args[2].value == [date(2024, 1, 1)]
        assert literal.dependencies == {start.id, ends.id}

    def test_an_axised_operand_is_refused(self):
        region = mo.Variable(["N", "S"], display_name="region")
        by_region = mo.Variable([date(2024, 1, 1), date(2024, 2, 1)], indexed_by=[region])
        with pytest.raises(ValueError, match="NETWORKDAYS"):
            NETWORKDAYS(by_region, date(2024, 3, 1))
        with pytest.raises(ValueError, match="NETWORKDAYS"):
            NETWORKDAYS(date(2024, 1, 1), date(2024, 3, 1), by_region)


class TestWorkday:
    def test_the_documented_examples(self):
        assert WORKDAY(WD_START, WD_DAYS).value == date(2009, 4, 30)
        assert WORKDAY(WD_START, WD_DAYS, WD_HOLIDAYS).value == date(2009, 5, 5)
        assert WORKDAY(V(WD_START), V(WD_DAYS), V(WD_HOLIDAYS)).value == date(2009, 5, 5)

    def test_steps_over_weekends_and_holidays_in_either_direction(self):
        fri = date(2024, 1, 5)
        assert WORKDAY(fri, 1).value == date(2024, 1, 8)
        assert WORKDAY(fri, 1, [date(2024, 1, 8)]).value == date(2024, 1, 9)
        assert WORKDAY(fri, 5).value == date(2024, 1, 12)
        assert WORKDAY(date(2024, 1, 8), -1).value == date(2024, 1, 5)
        assert WORKDAY(date(2024, 1, 8), -1, [date(2024, 1, 5)]).value == date(2024, 1, 4)
        assert WORKDAY(date(2024, 1, 6), 1).value == date(2024, 1, 8)      # from a Saturday
        assert WORKDAY(date(2024, 1, 7), -1).value == date(2024, 1, 5)     # from a Sunday
        assert WORKDAY(fri, 1, [date(2024, 1, 6), date(2024, 1, 7)]).value == date(2024, 1, 8)

    def test_zero_days_is_the_start_itself_even_on_a_weekend(self):
        assert WORKDAY(date(2024, 1, 6), 0).value == date(2024, 1, 6)
        assert WORKDAY(date(2024, 1, 6), 0, [date(2024, 1, 6)]).value == date(2024, 1, 6)

    def test_days_are_truncated_toward_zero(self):
        fri = date(2024, 1, 5)
        assert WORKDAY(fri, 1.9).value == date(2024, 1, 8)
        assert WORKDAY(fri, -1.9).value == date(2024, 1, 4)
        assert WORKDAY(fri, 0.5).value == fri
        with pytest.raises(TypeError, match="WORKDAY.*#VALUE!"):
            WORKDAY(fri, "ten")
        with pytest.raises(TypeError, match="WORKDAY needs a number"):
            WORKDAY(fri, V(None))

    def test_the_result_is_a_date(self):
        out = WORKDAY(V(date(2024, 1, 5), display_name="start"), V(3, unit="day"))
        assert out.value == date(2024, 1, 10) and type(out.value) is date
        assert out.value_type == 'datetime' and out._unit is None
        assert WORKDAY(datetime(2024, 1, 5, 15), 1).value == date(2024, 1, 8)

    def test_element_wise_over_the_start_or_the_days(self):
        starts = V([date(2024, 1, 5), date(2024, 1, 8), date(2024, 1, 9)])
        assert WORKDAY(starts, 1).value == [date(2024, 1, 8), date(2024, 1, 9), date(2024, 1, 10)]
        assert WORKDAY(date(2024, 1, 5), [1, 2, 3]).value == \
            [date(2024, 1, 8), date(2024, 1, 9), date(2024, 1, 10)]
        assert WORKDAY(starts, V([1, 2, 3]), [date(2024, 1, 9)]).value == \
            [date(2024, 1, 8), date(2024, 1, 11), date(2024, 1, 12)]
        assert WORKDAY(V([date(2024, 1, 5), "#N/A"]), 1).value == [date(2024, 1, 8), "#N/A"]
        assert WORKDAY(date(2024, 1, 5), V([1, "#DIV/0!"])).value == [date(2024, 1, 8), "#DIV/0!"]

    def test_the_holidays_are_one_operand_of_the_call_not_a_series(self):
        start = V(date(2024, 1, 5), display_name="start")
        hol = V([date(2024, 1, 8), date(2024, 1, 9)], display_name="hol")
        out = WORKDAY(start, 1, hol)
        assert out.value == date(2024, 1, 10)
        assert out.formula == "WORKDAY(start, 1, hol)"
        expr = out._expr
        assert isinstance(expr, FuncCall) and expr.func == "WORKDAY"
        assert isinstance(expr.args[2], VarRef) and expr.args[2].var is hol
        assert out.dependencies == {start.id, hol.id}
        assert WORKDAY(start, 1).formula == "WORKDAY(start, 1)"


# ══════════════════════════════════════════════════════════════════════
# The registry and a loop container
# ══════════════════════════════════════════════════════════════════════


class TestTheRegistry:
    @pytest.mark.parametrize("name", ["DATEDIF", "WEEKDAY", "WEEKNUM"])
    def test_one_call_per_period_under_a_bare_classic_name(self, name):
        assert is_elementwise(name) and not is_range_taking(name)
        assert spec(name).passthrough and not spec(name).bare
        assert file_prefix(name) == "" and excel_name(name) == name
        assert admits_in_loop(name, 1) and admits_in_loop(name, 3)

    @pytest.mark.parametrize("name", ["NETWORKDAYS", "WORKDAY"])
    def test_the_holidays_are_the_range_position(self, name):
        assert spec(name).shape == "range" and is_range_taking(name)
        assert range_positions(name) == (2,)
        assert not is_elementwise(name) and not admits_in_loop(name, 3)
        assert spec(name).passthrough and not spec(name).bare
        assert file_prefix(name) == "" and excel_name(name) == name


class TestOtherRenderers:
    """Excel-native calls: the JSON renderer keeps the structure and the
    backend hint, every other renderer inlines the value."""

    @pytest.mark.parametrize("build, name", [
        (lambda d: DATEDIF(d, date(2024, 3, 1), "D"), "DATEDIF"),
        (lambda d: WEEKDAY(d), "WEEKDAY"),
        (lambda d: WEEKNUM(d), "WEEKNUM"),
        (lambda d: NETWORKDAYS(d, date(2024, 3, 1), [date(2024, 1, 1)]), "NETWORKDAYS"),
        (lambda d: WORKDAY(d, 5), "WORKDAY"),
    ])
    def test_the_json_renderer_keeps_the_call_and_its_hint(self, build, name):
        out = mo.to_json(build(mo.Variable(date(2024, 1, 15)))._expr)
        assert out["kind"] == "funccall" and out["func"] == name
        assert out["render_backends"] == ["excel"]


class TestInsideALoop:
    def test_an_element_wise_function_runs_period_by_period(self):
        m = mo.Model("m", default_grain="month", default_start="2024-01", default_periods=3)
        with m:
            m.i = mo.MultiVariable("I")
            m.i.start = mo.Variable(date(2024, 1, 15))        # a Monday
            m.d = mo.MultiVariable("D", loop=True)
            with m.d as d:
                d.k = mo.lag(d.next, fill=0.0)
                d.due = m.i.start + d.k * 7
                d.w = WEEKDAY(d.due, 2)
                d.wk = WEEKNUM(d.due, 2)
                d.gap = DATEDIF(m.i.start, d.due, "D")
                d.next = d.k + d.w
        assert d.k.value == [0.0, 1.0, 2.0]
        assert d.due.value == [date(2024, 1, 15), date(2024, 1, 22), date(2024, 1, 29)]
        assert d.w.value == [1, 1, 1]
        assert d.wk.value == [3, 4, 5]
        assert d.gap.value == [0, 7, 14]
        assert d.next.value == [1.0, 2.0, 3.0]

    def test_a_holidays_range_is_refused_and_points_past_the_loop(self):
        m = mo.Model("m", default_grain="month", default_start="2024-01", default_periods=3)
        with pytest.raises(ValueError, match="NETWORKDAYS reads a whole series"):
            with m:
                m.i = mo.MultiVariable("I")
                m.i.start = mo.Variable(date(2024, 1, 15))
                m.d = mo.MultiVariable("D", loop=True)
                with m.d as d:
                    d.k = mo.lag(d.next, fill=0.0)
                    d.due = m.i.start + d.k * 7
                    d.n = NETWORKDAYS(m.i.start, d.due)
                    d.next = d.k + d.n
