# SPDX-License-Identifier: Apache-2.0
"""The rounding family, the powers and logarithms, ``SIGN`` and
``PRODUCT`` compute what the spreadsheet computes, with the formula as
the product.

``ROUNDUP`` / ``ROUNDDOWN`` / ``TRUNC`` move away from or toward zero on
the number as a cell shows it; ``CEILING`` / ``FLOOR`` keep the classic
function's sign rules; ``POWER`` / ``SQRT`` / ``EXP`` / ``LN`` / ``LOG``
/ ``LOG10`` give the spreadsheet's errors outside their domain; ``SIGN``
is 1, 0 or -1; ``PRODUCT`` reads a range as the sheet does. Each of the
first twelve is one call per period through the shared helper, so a
blank is 0, a date its serial number, numeric text its number and other
text ``#VALUE!``, an error cell passes through, scalars broadcast, tracks
lift, a rounded amount keeps its unit, and a number no cell can hold is
``#NUM!``.
"""

from __future__ import annotations

import datetime as dt
import math
import time

import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall, Literal, VarRef
from modeleon.functions.mathfn import (
    CEILING,
    EXP,
    FLOOR,
    LN,
    LOG,
    LOG10,
    POWER,
    PRODUCT,
    ROUNDDOWN,
    ROUNDUP,
    SIGN,
    SQRT,
    TRUNC,
)
from modeleon.functions.registry import (
    admits_in_loop,
    excel_name,
    file_prefix,
    is_elementwise,
    is_range_taking,
    spec,
)

ELEMENTWISE = {
    "ROUNDUP": ROUNDUP, "ROUNDDOWN": ROUNDDOWN, "TRUNC": TRUNC, "CEILING": CEILING,
    "FLOOR": FLOOR, "POWER": POWER, "SQRT": SQRT, "EXP": EXP, "LN": LN, "LOG": LOG,
    "LOG10": LOG10, "SIGN": SIGN,
}


# ══════════════════════════════════════════════════════════════════════
# The rounding family
# ══════════════════════════════════════════════════════════════════════


class TestRoundUp:
    @pytest.mark.parametrize("value, digits, expected", [
        (3.2, 0, 4.0),
        (76.9, 0, 77.0),
        (3.14159, 3, 3.142),
        (-3.14159, 1, -3.2),
        (31415.92654, -2, 31500.0),
        (-3.2, 0, -4.0),
        (2.5, 0, 3.0),
        (0.0, 0, 0.0),
        (1.001, 2, 1.01),
        (-1.001, 2, -1.01),
        (1e-20, 2, 0.01),
    ])
    def test_away_from_zero(self, value, digits, expected):
        assert ROUNDUP(value, digits).value == expected

    def test_on_the_number_as_shown(self):
        # 0.1 + 0.2 is held as 0.30000000000000004 and shown as 0.3.
        assert ROUNDUP(0.1 + 0.2, 1).value == 0.3
        assert ROUNDUP(mo.Variable(0.1) * 3, 1).value == 0.3
        assert ROUNDUP(2.0000000000000004, 0).value == 2.0

    def test_whole_numbers_stay_whole(self):
        assert ROUNDUP(7, 2).value == 7 and isinstance(ROUNDUP(7, 2).value, int)
        assert ROUNDUP(1234, -2).value == 1300 and isinstance(ROUNDUP(1234, -2).value, int)

    def test_a_list_and_digits_that_vary_by_period(self):
        v = mo.Variable([2.345, -2.345, 1250.0])
        assert ROUNDUP(v, 1).value == [2.4, -2.4, 1250.0]
        assert ROUNDUP(v, [2, 0, -2]).value == [2.35, -3.0, 1300.0]
        assert ROUNDUP(v, mo.Variable(1)).value == [2.4, -2.4, 1250.0]


class TestRoundDown:
    @pytest.mark.parametrize("value, digits, expected", [
        (3.2, 0, 3.0),
        (76.9, 0, 76.0),
        (3.14159, 3, 3.141),
        (-3.14159, 1, -3.1),
        (31415.92654, -2, 31400.0),
        (-3.9, 0, -3.0),
        (2.5, 0, 2.0),
        (1.999, 2, 1.99),
        (0.0, 0, 0.0),
    ])
    def test_toward_zero(self, value, digits, expected):
        assert ROUNDDOWN(value, digits).value == expected

    def test_on_the_number_as_shown(self):
        # 4.35 * 100 is held as 434.99999999999994 and shown as 435.
        assert ROUNDDOWN(4.35 * 100, 0).value == 435.0
        assert ROUNDDOWN(mo.Variable(4.35) * 100, 0).value == 435.0
        assert ROUNDDOWN(2.675 * 100, 0).value == 267.0

    def test_whole_numbers_stay_whole(self):
        assert ROUNDDOWN(1299, -2).value == 1200
        assert isinstance(ROUNDDOWN(1299, -2).value, int)


class TestTrunc:
    @pytest.mark.parametrize("value, digits, expected", [
        (8.9, None, 8.0),
        (-8.9, None, -8.0),
        (0.45, None, 0.0),
        (1234.5678, 2, 1234.56),
        (-1234.5678, 2, -1234.56),
        (1234.5678, -2, 1200.0),
        (-1234.5678, -2, -1200.0),
        (7, None, 7),
    ])
    def test_drops_the_decimals_toward_zero(self, value, digits, expected):
        out = TRUNC(value) if digits is None else TRUNC(value, digits)
        assert out.value == expected

    def test_on_the_number_as_shown(self):
        assert TRUNC(mo.Variable(4.35) * 100).value == 435.0

    def test_the_second_argument_is_written_only_when_given(self):
        x = mo.Variable(8.9, display_name="x")
        assert TRUNC(x).formula == "TRUNC(x)"
        assert TRUNC(x, 2).formula == "TRUNC(x, 2)"
        assert len(TRUNC(x)._expr.args) == 1 and len(TRUNC(x, 2)._expr.args) == 2


class TestCeiling:
    @pytest.mark.parametrize("value, significance, expected", [
        (2.5, 1, 3.0),
        (-2.5, -2, -4.0),
        (-2.5, 2, -2.0),
        (1.5, 0.1, 1.5),
        (0.234, 0.01, 0.24),
        (4.2, 2, 6.0),
        (4.0, 2, 4.0),
        (0.0, 0.5, 0.0),
        (-0.5, 2, 0.0),
        (7, 2, 8),
        (6, 2, 6),
        (-7, 2, -6),
        (-7, -2, -8),
    ])
    def test_away_from_zero_to_a_multiple(self, value, significance, expected):
        assert CEILING(value, significance).value == expected

    def test_a_positive_number_with_a_negative_significance_is_num(self):
        assert CEILING(2.5, -2).value == '#NUM!'
        assert CEILING(mo.Variable([2.5, -2.5]), -2).value == ['#NUM!', -4.0]

    def test_a_zero_significance_gives_zero(self):
        assert CEILING(5, 0).value == 0
        assert CEILING(mo.Variable([5.0, -5.0]), 0).value == [0, 0]

    def test_the_quotient_is_read_as_shown(self):
        # 1.1 / 0.1 is held as 11.000000000000002 and shown as 11.
        assert CEILING(1.1, 0.1).value == 1.1
        assert CEILING(mo.Variable(1.1), 0.1).value == 1.1
        assert CEILING(mo.Variable(0.7) * 3, 0.1).value == 2.1

    def test_a_quotient_of_sixteen_whole_digits_is_taken_as_held(self):
        # No decimals are left to show, so an exact multiple stays.
        assert CEILING(1000000000000001, 1).value == 1000000000000001
        assert CEILING(1000000000000000.5, 1).value == 1000000000000001.0
        assert FLOOR(1000000000000000.5, 1).value == 1000000000000000.0
        assert CEILING(51234567890123.45, 0.01).value == 51234567890123.45
        assert FLOOR(51234567890123.45, 0.01).value == 51234567890123.45


class TestFloor:
    @pytest.mark.parametrize("value, significance, expected", [
        (3.7, 2, 2.0),
        (-2.5, -2, -2.0),
        (-2.5, 2, -4.0),
        (1.58, 0.1, 1.5),
        (0.234, 0.01, 0.23),
        (4.0, 2, 4.0),
        (0.0, 0.5, 0.0),
        (7, 2, 6),
        (6, 2, 6),
        (-7, 2, -8),
        (-7, -2, -6),
    ])
    def test_toward_zero_to_a_multiple(self, value, significance, expected):
        assert FLOOR(value, significance).value == expected

    def test_a_positive_number_with_a_negative_significance_is_num(self):
        assert FLOOR(2.5, -2).value == '#NUM!'

    def test_a_zero_significance_divides_by_zero_unless_the_number_is_zero(self):
        assert FLOOR(5, 0).value == '#DIV/0!'
        assert FLOOR(mo.Variable([5.0, -5.0, 0.0]), 0).value == ['#DIV/0!', '#DIV/0!', 0.0]
        assert FLOOR(0, 0).value == 0 and isinstance(FLOOR(0, 0).value, int)
        assert FLOOR(mo.Variable([None]), mo.Variable([None])).value == [0]

    def test_the_quotient_is_read_as_shown(self):
        # 0.3 / 0.1 is held as 2.9999999999999996 and shown as 3.
        assert FLOOR(0.3, 0.1).value == 0.3
        assert FLOOR(mo.Variable(0.7) * 3, 0.1).value == 2.1


class TestANegativeZeroIsZero:
    @pytest.mark.parametrize("out", [
        ROUNDUP(-0.0, 0), ROUNDDOWN(-0.4, 0), TRUNC(-0.125), TRUNC(-0.004, 2),
        CEILING(-0.5, 2), FLOOR(-0.0, 1), POWER(-0.0, 3),
    ])
    def test_as_a_cell_shows_it(self, out):
        assert out.value == 0 and math.copysign(1.0, out.value) == 1.0


class TestANumberNoCellCanHoldIsNum:
    def test_rounded_to_a_multiple_past_the_largest_number(self):
        assert ROUNDUP(2.5, -320).value == '#NUM!'            # 1e320 fits no cell
        assert ROUNDUP(5, -320).value == '#NUM!'
        assert ROUNDDOWN(2.5, -320).value == 0 and TRUNC(-2.5, -320).value == 0
        assert ROUNDDOWN(2.5, 400).value == 2.5
        assert CEILING(1.7e308, 1e308).value == '#NUM!'
        assert ROUNDUP(math.inf, 0).value == '#NUM!'
        assert CEILING(math.nan, 1).value == '#NUM!'

    def test_a_quotient_past_the_largest_number_or_below_the_smallest(self):
        # The sheet divides in doubles: 1e300 / 1e-300 is past any cell,
        # 1e-300 / 1e300 is 0.
        assert CEILING(1e300, 1e-300).value == '#NUM!'
        assert FLOOR(1e300, 1e-300).value == '#NUM!'
        assert CEILING(-1e300, -1e-300).value == '#NUM!'
        assert CEILING(1e-300, 1e300).value == 0 and FLOOR(-1e-300, 1e300).value == 0
        assert FLOOR(1.7e308, 1e308).value == 1e308

    @pytest.mark.parametrize("fn", [SQRT, EXP, LN, LOG, LOG10, SIGN])
    def test_an_operand_no_cell_holds(self, fn):
        assert fn(math.inf).value == '#NUM!' and fn(math.nan).value == '#NUM!'

    def test_a_power_or_a_product_past_it(self):
        assert POWER(10, 400).value == '#NUM!' and POWER(2.0, 1024).value == '#NUM!'
        assert PRODUCT(1e200, 1e200).value == '#NUM!'
        assert PRODUCT(mo.Variable([10 ** 200, 10 ** 200])).value == '#NUM!'


# ══════════════════════════════════════════════════════════════════════
# The powers and logarithms
# ══════════════════════════════════════════════════════════════════════


class TestPower:
    def test_the_values(self):
        assert POWER(5, 2).value == 25 and isinstance(POWER(5, 2).value, int)
        assert POWER(98.6, 3.2).value == pytest.approx(2401077.222, abs=1e-3)
        assert POWER(4, 5 / 4).value == pytest.approx(5.656854, abs=1e-6)
        assert POWER(2, -1).value == 0.5
        assert POWER(-2, 3).value == -8
        assert POWER(0, 5).value == 0
        assert POWER(mo.Variable([2, 3, 4]), 2).value == [4, 9, 16]
        assert POWER(2, [1, 2, 3]).value == [2, 4, 8]

    def test_the_spreadsheet_s_errors(self):
        assert POWER(0, 0).value == '#NUM!'
        assert POWER(-8, 0.5).value == '#NUM!'
        assert POWER(0, -1).value == '#DIV/0!'
        assert POWER(10.0, 400).value == '#NUM!'
        assert POWER(10, 400).value == '#NUM!'

    def test_a_negative_base_has_a_real_odd_root(self):
        assert POWER(-8, 1 / 3).value == -2.0
        assert POWER(-32, 0.2).value == -2.0 and POWER(-8, -1 / 3).value == -0.5
        assert POWER(-27, 1 / 3).value == pytest.approx(-3.0)
        assert POWER(mo.Variable([-8.0, 8.0]), 1 / 3).value == [-2.0, 2.0]
        # The inverse of an odd whole number to within the 15th digit.
        assert POWER(-8, 0.333333333333333).value == pytest.approx(-2.0)
        assert POWER(-8, 0.33333333333333).value == '#NUM!'

    @pytest.mark.parametrize("exponent", [0.5, 0.4, 2 / 3, 1.5, -2.5, 1e-16, 5e-324])
    def test_any_other_fraction_of_a_negative_base_is_num(self, exponent):
        assert POWER(-8, exponent).value == '#NUM!'

    def test_a_whole_power_past_any_cell_is_num_without_computing_it(self):
        started = time.perf_counter()
        assert POWER(3, 10 ** 7).value == '#NUM!'          # some seconds to compute in full
        assert POWER(-2, 10 ** 7 + 1).value == '#NUM!'
        assert time.perf_counter() - started < 1.0
        assert POWER(2, 1023).value == 2 ** 1023 and POWER(2, 1024).value == '#NUM!'
        assert POWER(-1, 10 ** 9 + 1).value == -1 and POWER(1, 10 ** 9).value == 1

    def test_a_power_too_small_for_a_cell_is_zero(self):
        assert POWER(8, -400).value == 0.0 and POWER(0.5, 1100).value == 0.0

    def test_the_unit_follows_the_operator(self):
        amount = mo.Variable([3.0, 8.0], unit="$")
        assert str(POWER(amount, 2)._unit) == str((amount ** 2)._unit)
        assert POWER(amount, 0.5)._unit is None
        assert POWER(amount, mo.Variable(2))._unit is None
        assert POWER(2, amount)._unit is None


class TestRootsExponentialsAndLogarithms:
    def test_sqrt(self):
        assert SQRT(16).value == 4.0
        assert SQRT(2).value == pytest.approx(1.414214, abs=1e-6)
        assert SQRT(0).value == 0.0
        assert SQRT(-16).value == '#NUM!'
        assert SQRT(mo.Variable([16.0, -16.0, 2.25])).value == [4.0, '#NUM!', 1.5]

    def test_exp(self):
        assert EXP(0).value == 1.0
        assert EXP(1).value == pytest.approx(2.718282, abs=1e-6)
        assert EXP(-1).value == pytest.approx(0.367879, abs=1e-6)
        assert EXP(1000).value == '#NUM!'

    def test_ln(self):
        assert LN(86).value == pytest.approx(4.454347, abs=1e-6)
        assert LN(math.e).value == 1.0
        assert LN(1).value == 0.0
        assert LN(0).value == '#NUM!' and LN(-1).value == '#NUM!'

    def test_log(self):
        assert LOG(10).value == 1.0 and LOG(100).value == 2.0
        assert LOG(8, 2).value == 3.0
        assert LOG(1000, 10).value == 3.0             # exact, as a cell shows it
        assert LOG(81, 3).value == 4.0
        assert LOG(86, 2.7182818).value == pytest.approx(4.454347, abs=1e-6)
        assert LOG(8, 1).value == '#DIV/0!'
        assert LOG(-1).value == '#NUM!' and LOG(0).value == '#NUM!'
        assert LOG(8, -2).value == '#NUM!' and LOG(8, 0).value == '#NUM!'

    def test_log10(self):
        assert LOG10(86).value == pytest.approx(1.934498, abs=1e-6)
        assert LOG10(10).value == 1.0 and LOG10(1e5).value == 5.0
        assert LOG10(0).value == '#NUM!'

    def test_the_base_is_written_only_when_given(self):
        x = mo.Variable(8.0, display_name="x")
        assert LOG(x).formula == "LOG(x)" and LOG(x, 2).formula == "LOG(x, 2)"
        assert len(LOG(x)._expr.args) == 1 and len(LOG(x, 2)._expr.args) == 2

    @pytest.mark.parametrize("fn", [SQRT, EXP, LN, LOG, LOG10])
    def test_a_float_without_the_operand_s_unit(self, fn):
        amount = mo.Variable([3.0, 8.0], unit="$")
        out = fn(amount)
        assert out.value_type == 'float' and out._unit is None
        assert fn(mo.Variable([3, 8])).value_type == 'float'


class TestSign:
    def test_the_values(self):
        assert SIGN(10).value == 1 and SIGN(0).value == 0 and SIGN(-0.00001).value == -1
        assert SIGN(mo.Variable(4) - 4).value == 0
        assert SIGN(mo.Variable([3.0, -8.0, 0.0])).value == [1, -1, 0]

    def test_a_whole_number_without_the_operand_s_unit(self):
        out = SIGN(mo.Variable([3.0, -8.0], unit="$"))
        assert out.value_type == 'int' and out._unit is None


# ══════════════════════════════════════════════════════════════════════
# The shared law: one call per period
# ══════════════════════════════════════════════════════════════════════


class TestOneCallPerPeriod:
    def test_a_blank_is_zero(self):
        blank = mo.Variable([None])
        assert ROUNDUP(blank, 0).value == [0]
        assert TRUNC(blank).value == [0]
        assert CEILING(5, blank).value == [0]
        assert POWER(blank, 2).value == [0] and POWER(2, blank).value == [1]
        assert SQRT(blank).value == [0.0]
        assert EXP(blank).value == [1.0]
        assert LN(blank).value == ['#NUM!']
        assert SIGN(blank).value == [0]

    def test_numeric_text_is_its_number_and_other_text_is_value(self):
        words = mo.Variable(["abc", "4", "2.5", "50%"])
        assert ROUNDUP(words, 0).value == ['#VALUE!', 4, 3.0, 1.0]
        assert CEILING(words, 1).value == ['#VALUE!', 4, 3.0, 1.0]
        assert POWER(2, words).value == [
            '#VALUE!', 16, pytest.approx(5.656854, abs=1e-6), pytest.approx(1.414214, abs=1e-6),
        ]
        assert SIGN(words).value == ['#VALUE!', 1, 1, 1]
        assert SQRT("16").value == 4.0 and SQRT("abc").value == '#VALUE!'
        assert ROUNDUP("1,234.5", 0).value == 1235            # read as VALUE reads it

    def test_a_date_is_its_serial_number(self):
        m = mo.Model("m", default_grain="month", default_start="2024-01", default_periods=3)
        with m:
            m.when = mo.EOMONTH(mo.time.start, 0)         # 45322, 45351, 45382
            m.week = FLOOR(m.when, 7)
            m.week_end = CEILING(mo.time.end, 7)
            m.sign = SIGN(m.when)
        assert m.week.value == [45318, 45346, 45381]
        assert m.week_end.value == [45325, 45353, 45388]
        assert m.sign.value == [1, 1, 1]
        day = mo.Variable(dt.date(2024, 2, 29))
        noon = mo.Variable(dt.datetime(2024, 2, 29, 12, 0))
        assert ROUNDUP(day, -1).value == 45360 and TRUNC(day).value == 45351
        assert TRUNC(noon).value == 45351 and ROUNDDOWN(noon, 1).value == 45351.5
        assert SQRT(day).value == pytest.approx(212.957742287056)
        assert FLOOR(day, 7).value_type == 'float'

    def test_a_boolean_is_one_or_zero(self):
        assert ROUNDUP(True, 0).value == 1 and ROUNDUP(False, 0).value == 0
        assert SIGN(True).value == 1 and SQRT(True).value == 1.0
        assert POWER(True, 2).value == 1 and CEILING(True, 2).value == 2

    def test_an_error_cell_is_the_period_s_result(self):
        e = mo.Variable([2.0, 2.0]) / mo.Variable([1.0, 0.0])
        assert e.value == [2.0, '#DIV/0!']
        assert ROUNDUP(e, 0).value == [2.0, '#DIV/0!']
        assert SQRT(e).value == [pytest.approx(1.414214, abs=1e-6), '#DIV/0!']
        assert POWER(e, 2).value == [4.0, '#DIV/0!']
        assert POWER(2, e).value == [4.0, '#DIV/0!']
        assert CEILING(1.0, e).value == [2.0, '#DIV/0!']
        assert SIGN(e).value == [1, '#DIV/0!']

    def test_scalars_only_is_one_number(self):
        out = ROUNDUP(3.3, 0)
        assert out.value == 4.0 and out.var_type == 'scalar' and out.value_type == 'float'
        assert out.formula == "ROUNDUP(3.3, 0)"
        assert CEILING(mo.Variable(2.5), mo.Variable(2)).var_type == 'scalar'

    def test_scalars_and_a_one_value_series_broadcast(self):
        a = mo.Variable([2.345, -2.345, 1250.0])
        assert ROUNDUP(a, mo.Variable(1)).value == [2.4, -2.4, 1250.0]
        assert ROUNDUP(a, mo.Variable([1])).value == [2.4, -2.4, 1250.0]
        assert POWER(mo.Variable([2.0]), [1, 2, 3]).value == [2.0, 4.0, 8.0]
        assert ROUNDUP(a, 0).var_type == 'list'

    def test_series_of_different_lengths_are_refused_as_the_operators_refuse_them(self):
        with pytest.raises(ValueError, match="different lengths"):
            ROUNDUP(mo.Variable([1.1, 2.2, 3.3]), mo.Variable([1, 2]))

    def test_tracked_operands_lift_per_track(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"),
                     default_grain="month", default_start="2026-01", default_periods=2)
        with m:
            m.x = mo.Variable(plan=[1.26, 2.0], actual=[1.0, None])
            m.r = ROUNDUP(m.x, 1)
            m.s = SQRT(m.x)
        assert dict(m.r._value.items()) == {"plan": [1.3, 2.0], "actual": [1.0, 0]}
        assert dict(m.s._value.items()) == {
            "plan": [pytest.approx(1.122497, abs=1e-6), pytest.approx(1.414214, abs=1e-6)],
            "actual": [1.0, 0.0],
        }
        assert m.r.var_type == 'list'

    def test_value_type_is_int_only_when_the_result_is_whole_by_construction(self):
        assert ROUNDUP(mo.Variable([3, 8]), -1).value_type == 'int'
        assert ROUNDUP(7, 2).value_type == 'int' and TRUNC(7).value_type == 'int'
        assert ROUNDUP(mo.Variable([3.0, 8.0]), 0).value_type == 'float'
        assert ROUNDUP(mo.Variable(["4"]), 0).value_type == 'float'
        assert CEILING(mo.Variable([3, 8]), 2).value_type == 'int'
        assert CEILING(mo.Variable([3, 8]), 0.5).value_type == 'float'
        assert CEILING(mo.Variable([3.0, 8.0]), mo.Variable(2)).value_type == 'float'
        assert isinstance(CEILING(5.0, 0).value, float) and isinstance(CEILING(5, 0).value, int)
        assert POWER(mo.Variable([3, 8]), 2).value_type == 'int'
        assert POWER(mo.Variable([3, 8]), -1).value_type == 'float'
        assert POWER(mo.Variable([3, 8]), mo.Variable(2)).value_type == 'float'

    def test_a_rounded_amount_keeps_its_unit(self):
        a = mo.Variable([3.0, 8.0], unit="$")
        step = mo.Variable(100.0, unit="$")
        hours = mo.Variable(2.0, unit="hr")
        for fn in (ROUNDUP, ROUNDDOWN, TRUNC, CEILING, FLOOR):
            assert str(fn(a, 1)._unit) == "$", fn
        assert str(CEILING(a, step)._unit) == "$"
        # A multiple is in the unit of the step: a step in hours is refused.
        with pytest.raises(ValueError, match="mo.CEILING takes its values in one unit"):
            CEILING(a, hours)

    def test_the_formula_names_the_operands_and_the_ast_is_a_call(self):
        a = mo.Variable([3.0, 8.0], display_name="a")
        k = mo.Variable(2, display_name="k")
        assert ROUNDUP(a, 1).formula == "ROUNDUP(a, 1)"
        assert CEILING(a, k).formula == "CEILING(a, k)"
        assert POWER(a, k).formula == "POWER(a, k)"
        assert SIGN(a).formula == "SIGN(a)"
        expr = ROUNDUP(a, 1)._expr
        assert isinstance(expr, FuncCall) and expr.func == "ROUNDUP"
        assert isinstance(expr.args[0], VarRef) and expr.args[0].var is a
        assert isinstance(expr.args[1], Literal) and expr.args[1].value == 1

    @pytest.mark.parametrize("name", sorted(ELEMENTWISE))
    def test_an_axised_operand_is_refused(self, name):
        region = mo.Variable(["N", "S"], display_name="region")
        by_region = mo.Variable([1.0, 2.0], indexed_by=[region])
        fn = ELEMENTWISE[name]
        with pytest.raises(ValueError, match=name):
            fn(by_region, 1) if name in ("ROUNDUP", "ROUNDDOWN", "CEILING", "FLOOR", "POWER") \
                else fn(by_region)


# ══════════════════════════════════════════════════════════════════════
# PRODUCT, a reduction
# ══════════════════════════════════════════════════════════════════════


class TestProduct:
    def test_one_row_reduces_to_one_number(self):
        out = PRODUCT(mo.Variable([5.0, 15.0, 30.0]))
        assert out.value == 2250.0 and out.var_type == 'scalar'
        assert PRODUCT([5, 15, 30]).value == 2250

    def test_variadic_like_sum(self):
        assert PRODUCT(5, 15, 30).value == 2250
        assert PRODUCT(5, 15, 30).formula == "PRODUCT(5, 15, 30)"
        assert PRODUCT(mo.Variable(2.0), mo.Variable(3.0)).value == 6.0
        # A series beside a scalar: the scalar multiplies every period, as
        # ``a * 2`` does; the row's product times 2 is
        # ``PRODUCT(row) * 2``.
        assert PRODUCT(mo.Variable([5.0, 15.0, 30.0]), 2).value == [10.0, 30.0, 60.0]
        assert (PRODUCT(mo.Variable([5.0, 15.0, 30.0])) * 2).value == 4500.0

    def test_two_series_are_multiplied_period_by_period(self):
        # Several series are combined period by period, as ``a * b`` combines
        # them; the product of all their cells is two products.
        a = mo.Variable([2.0, 3.0])
        b = mo.Variable([4.0, 5.0])
        out = PRODUCT(a, b)
        assert out.value == [8.0, 15.0] and out.var_type == 'list'
        assert (PRODUCT(a) * PRODUCT(b)).value == 120.0

    def test_the_range_is_read_as_the_sheet_reads_it(self):
        assert PRODUCT(mo.Variable([2.0, None, 3.0])).value == 6.0
        assert PRODUCT(mo.Variable([None, None])).value == 0
        assert PRODUCT(mo.Variable(["#N/A", "#DIV/0!", 3.0])).value == "#N/A"
        # Period by period: each period's first error, a blank skipped.
        assert PRODUCT(mo.Variable([1.0, 2.0, None]),
                       mo.Variable(["#VALUE!", 1.0, None])).value == ["#VALUE!", 2.0, 0]

    def test_no_argument_is_refused(self):
        with pytest.raises(ValueError, match=r"PRODUCT\(\) needs at least one argument"):
            PRODUCT()

    def test_text_is_refused_at_build_time_as_sum_refuses_it(self):
        with pytest.raises(TypeError, match="PRODUCT multiplies numbers; 'abc' is not one"):
            PRODUCT(mo.Variable(["abc"]))
        with pytest.raises(TypeError, match="give it the number"):
            PRODUCT(mo.Variable([2.0, "abc"]), 3)

    def test_a_boolean_is_one_or_zero_as_libreoffice_reads_it(self):
        assert PRODUCT(True, 2.5).value == 2.5 and PRODUCT(False, 2.5).value == 0
        assert PRODUCT(mo.Variable([True, False, True])).value == 0

    def test_a_date_is_its_serial_number(self):
        assert PRODUCT(mo.Variable(dt.date(2024, 2, 29)), 1).value == 45351
        assert PRODUCT(mo.Variable([dt.date(2024, 1, 31), 2])).value == 90644

    def test_a_line_with_tracks_reduces_track_by_track(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"),
                     default_grain="month", default_start="2026-01", default_periods=2)
        with m:
            m.x = mo.Variable(plan=[1.5, 2.0], actual=[3.0, None])
            m.p = PRODUCT(m.x)
        assert dict(m.p._value.items()) == {"plan": 3.0, "actual": 3.0}
        assert m.p.formula == "PRODUCT(x)"

    def test_the_units_multiply_and_a_whole_number_only_over_whole_numbers(self):
        a = mo.Variable([2.0, 3.0], unit="$")
        price, volume = mo.Variable(2.0, unit="$/t"), mo.Variable(3.0, unit="t")
        assert str(PRODUCT(price, volume)._unit) == "$"
        # A row in a dimension is refused: its cells' product is $ to the
        # power of their count; a row of pure numbers multiplies.
        with pytest.raises(ValueError, match="a row in \\$"):
            PRODUCT(a)
        assert PRODUCT(mo.Variable([1.1, 1.2])).value == pytest.approx(1.32)
        assert PRODUCT(mo.Variable([2.0, 3.0])).value_type == 'float'
        assert PRODUCT(mo.Variable([2, 3])).value_type == 'int'
        assert PRODUCT(mo.Variable([2, 3]), 3).value_type == 'int'
        assert PRODUCT(mo.Variable([2, 3]), 2.5).value_type == 'float'
        assert PRODUCT(mo.Variable(dt.date(2024, 2, 29)), 2).value_type == 'float'

    def test_an_axised_operand_is_refused(self):
        region = mo.Variable(["N", "S"], display_name="region")
        by_region = mo.Variable([1.0, 2.0], indexed_by=[region])
        with pytest.raises(ValueError, match="PRODUCT"):
            PRODUCT(by_region)


# ══════════════════════════════════════════════════════════════════════
# The registry, a recurrence template, a loop container
# ══════════════════════════════════════════════════════════════════════


class TestTheRegistry:
    @pytest.mark.parametrize("name", sorted(ELEMENTWISE))
    def test_each_is_element_wise_with_a_bare_classic_name(self, name):
        assert is_elementwise(name) and not is_range_taking(name)
        assert spec(name).passthrough and not spec(name).bare
        assert file_prefix(name) == "" and excel_name(name) == name
        assert admits_in_loop(name, 2) and admits_in_loop(name, 1)

    def test_product_is_a_reduction_and_one_call_per_period_over_two_values(self):
        assert spec("PRODUCT").shape == "reduction" and is_range_taking("PRODUCT")
        assert not is_elementwise("PRODUCT") and not admits_in_loop("PRODUCT", 1)
        assert spec("PRODUCT").per_period_arity == 2
        assert admits_in_loop("PRODUCT", 2) and admits_in_loop("PRODUCT", 3)


class TestInARecurrenceTemplate:
    def test_a_template_computes_what_the_cell_computes(self):
        assert mo.recurrence(100.0, "ROUNDUP({prev} * 1.1, 0)", periods=4).value == \
            [100.0, 110.0, 121.0, 134.0]
        assert mo.recurrence(100.0, "CEILING({prev} * 1.07, 5)", periods=3).value == \
            [100.0, 110.0, 120.0]
        assert mo.recurrence(100.0, "FLOOR({prev} * 0.9, 1)", periods=3).value == \
            [100.0, 90.0, 81.0]
        assert mo.recurrence(256.0, "SQRT({prev})", periods=3).value == [256.0, 16.0, 4.0]
        assert mo.recurrence(2, "POWER({prev}, 2)", periods=3).value == [2, 4, 16]
        assert mo.recurrence(5.5, "TRUNC({prev} * 1.5)", periods=3).value == [5.5, 8.0, 12.0]
        assert mo.recurrence(1.0, "SIGN({prev}) * ROUNDDOWN({prev} * 2.5, 0)", periods=3).value == \
            [1.0, 2.0, 5.0]
        grown = mo.recurrence(1.0, "EXP(LN({prev}) + LOG(100, 10) + LOG10(10))", periods=2)
        assert grown.value == [1.0, pytest.approx(math.e ** 3)]

    def test_a_number_no_cell_holds_is_num_and_carries_on(self):
        assert mo.recurrence(1.0, "EXP({prev} * 1000)", periods=3).value == [1.0, '#NUM!', '#NUM!']
        assert mo.recurrence(1e308, "ROUNDUP({prev} * 10, 0)", periods=3).value == \
            [1e308, '#NUM!', '#NUM!']
        assert mo.recurrence(-8.0, "POWER({prev}, 1 / 3)", periods=2).value == [-8.0, -2.0]


class TestInsideALoop:
    def test_an_element_wise_function_runs_period_by_period(self):
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
        with m:
            m.d = mo.MultiVariable("D", loop=True)
            with m.d as d:
                d.bal = mo.lag(d.r, fill=10.0) * 1.37
                d.cap = mo.MIN(d.bal, 15.0)
                d.r = ROUNDUP(d.cap, 0)
                d.c = CEILING(d.bal, 5.0)
                d.s = SQRT(d.c)
        assert d.bal.value == pytest.approx([13.7, 19.18, 20.55])
        assert d.r.value == pytest.approx([14.0, 15.0, 15.0])
        assert d.c.value == pytest.approx([15.0, 20.0, 25.0])
        assert d.s.value == pytest.approx([15.0 ** 0.5, 20.0 ** 0.5, 5.0])

    def test_the_reduction_is_refused_and_points_past_the_loop(self):
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
        with pytest.raises(ValueError, match="PRODUCT reads a whole series"):
            with m:
                m.d = mo.MultiVariable("D", loop=True)
                with m.d as d:
                    d.bal = mo.lag(d.x, fill=10.0) * 1.37
                    d.total = PRODUCT(d.bal)
                    d.x = d.bal - d.total
