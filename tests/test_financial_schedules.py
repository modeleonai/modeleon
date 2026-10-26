# SPDX-License-Identifier: Apache-2.0
"""The loan schedule — ``NPER``, ``IPMT``, ``PPMT``, ``RATE``, ``CUMIPMT``,
``CUMPRINC`` — and depreciation — ``SLN``, ``DB``, ``DDB``, ``SYD`` —
compute what the spreadsheet computes, with the formula as the product.

The numbers are the spreadsheet's own: the documented examples, and a
recalculation of the fractional, zero-rate and negative cases. Each
function is one call per period through the shared helper, so a blank
is 0, numeric text its number and other text ``#VALUE!``, an error cell
passes through, scalars broadcast, tracks lift, an amount keeps its
unit, and a period-shaped operand — ``mo.time.index`` — makes the call
a schedule row.
"""

from __future__ import annotations

import math

import numpy as np
import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall, Literal, VarRef
from modeleon.functions import financial
from modeleon.functions.financial import (
    CUMIPMT,
    CUMPRINC,
    DB,
    DDB,
    IPMT,
    NPER,
    PMT,
    PPMT,
    PV,
    RATE,
    SLN,
    SYD,
)
from modeleon.functions.registry import (
    admits_in_loop,
    excel_name,
    file_prefix,
    is_elementwise,
    is_range_taking,
    spec,
)

NAMES = {
    "NPER": NPER, "IPMT": IPMT, "PPMT": PPMT, "RATE": RATE, "CUMIPMT": CUMIPMT,
    "CUMPRINC": CUMPRINC, "SLN": SLN, "DB": DB, "DDB": DDB, "SYD": SYD,
}

MONTHLY = 0.09 / 12


def close(value, expected, rel=1e-9):
    return value == pytest.approx(expected, rel=rel, abs=1e-9)


# ══════════════════════════════════════════════════════════════════════
# The loan schedule
# ══════════════════════════════════════════════════════════════════════


class TestNper:
    @pytest.mark.parametrize("args, expected", [
        ((0.01, -100, -1000, 10000, 1), 59.6738656742946),
        ((0.01, -100, -1000, 10000), 60.0821228537617),
        ((0.01, -100, -1000), -9.57859403981317),
        ((0.0, -100, 1000), 10.0),
        ((0.05, -100, 0, 0), 0.0),
        ((0.05, 0, 0), 0.0),
        ((0.1, -100, 1000, 0, 1), 25.1588579280968),
        ((0.1, -100, 1000, 0, 2), 25.1588579280968),      # any other type reads as 1
        ((0.05, 0, -1000, 2000), 14.2066990828226),          # no payment: a sum doubling
    ])
    def test_the_values(self, args, expected):
        assert close(NPER(*args).value, expected)

    def test_the_spreadsheet_s_errors(self):
        assert NPER(0, 0, 1000).value == '#NUM!'              # no payment at no rate
        assert NPER(0.1, 100, -1000).value == '#NUM!'         # the payment is the interest
        assert NPER(0.1, 50, -1000).value == '#NUM!'          # and does not even cover it
        # At -100% no number of periods pays anything off (LibreOffice
        # divides by log1p(-1) = -inf and shows 0).
        assert NPER(-1, -100, 1000).value == '#NUM!'
        assert NPER(-1.5, -100, 1000).value == '#NUM!'

    def test_no_unit_and_a_float(self):
        pmt = mo.Variable(-200.0, unit="$")
        loan = mo.Variable(8000.0, unit="$")
        out = NPER(0.01, pmt, loan)
        assert out._unit is None and out.value_type == 'float' and out.var_type == 'scalar'
        assert NPER(0.01, mo.Variable(-200), mo.Variable(8000)).value_type == 'float'

    def test_an_optional_argument_is_written_only_when_given(self):
        rate = mo.Variable(0.01, display_name="rate")
        pmt = mo.Variable(-100.0, display_name="pmt")
        pv = mo.Variable(-1000.0, display_name="pv")
        assert NPER(rate, pmt, pv).formula == "NPER(rate, pmt, pv)"
        assert NPER(rate, pmt, pv, 10000.0).formula == "NPER(rate, pmt, pv, 10000.0)"
        assert NPER(rate, pmt, pv, when=1).formula == "NPER(rate, pmt, pv, 0, 1)"
        assert len(NPER(rate, pmt, pv)._expr.args) == 3
        assert len(NPER(rate, pmt, pv, when=1)._expr.args) == 5


class TestIpmt:
    @pytest.mark.parametrize("args, expected", [
        ((0.1 / 12, 1, 36, 8000), -66.6666666666667),
        ((0.1, 3, 3, 8000), -292.447129909365),
        ((0.1, 1, 3, 8000, 0, 1), 0.0),
        ((0.1, 2, 3, 8000, 0, 1), -507.552870090635),
        ((0.1, 2, 3, 8000, 0, 2), -507.552870090635),
        ((0.0, 2, 10, 1000), 0.0),
        ((0.1, 2.5, 3, 8000), -428.544451994763),
        ((0.1, 2, 3.5, 8000), -597.961725204323),
        ((0.1, 2, 3, 8000, 1000), -528.09667673716),
        ((0.1, 1, 3, -8000), 800.0),
    ])
    def test_the_values(self, args, expected):
        assert close(IPMT(*args).value, expected)

    def test_a_zero_rate_is_a_plain_zero(self):
        value = IPMT(0.0, 2, 10, 1000).value
        assert value == 0.0 and math.copysign(1.0, value) == 1.0

    def test_the_spreadsheet_s_errors(self):
        assert IPMT(0.1, 0, 3, 8000).value == '#NUM!'
        assert IPMT(0.1, 4, 3, 8000).value == '#NUM!'
        assert IPMT(0.1, 1, 0, 8000).value == '#NUM!'
        assert IPMT(-1.5, 1, 3, 8000).value == '#NUM!'

    def test_a_rate_of_minus_100_percent_is_computed_as_the_spreadsheet_computes_it(self):
        # Every amount is gone after a period: the first period's interest
        # is the whole loan and the payment nothing (LibreOffice's values).
        assert IPMT(-1, 1, 3, 8000).value == 8000.0
        assert PPMT(-1, 1, 3, 8000).value == -8000.0
        assert IPMT(-1, 2, 3, 8000).value == 0.0
        assert PPMT(-1, 2, 3, 8000).value == 0.0
        assert IPMT(-1, 1, 3, 8000, 0, 1).value == 0.0

    def test_the_first_period_needs_no_payment(self):
        # (1 + rate) ** nper overflows on a term this long; the first
        # period's interest is the loan's own, as the spreadsheet gives it,
        # and the rest is #NUM! on both sides.
        assert IPMT(0.1, 1, 10000, 1000).value == -100.0
        assert IPMT(0.1, 1, 10000, 1000, 500).value == -100.0
        assert CUMIPMT(0.1, 10000, 1000, 1, 1, 0).value == -100.0
        assert IPMT(0.1, 2, 10000, 1000).value == '#NUM!'
        assert PPMT(0.1, 1, 10000, 1000).value == '#NUM!'

    def test_a_series_per_is_a_schedule_row(self):
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
        with m:
            m.rate = mo.Variable(0.1, display_name="rate")
            m.loan = mo.Variable(8000.0, display_name="loan", unit="$")
            m.interest = IPMT(m.rate, mo.time.index, 3, m.loan)
        assert m.interest.var_type == 'list' and m.interest.value_type == 'float'
        assert m.interest.value == [pytest.approx(IPMT(0.1, p, 3, 8000).value) for p in (1, 2, 3)]
        assert str(m.interest._unit) == "$"
        assert IPMT(0.1, [1, 2, 3], 3, 8000).value == m.interest.value


class TestPpmt:
    @pytest.mark.parametrize("args, expected", [
        ((0.1 / 12, 1, 24, 2000), -75.6231860083664),
        ((0.08, 10, 10, 200000), -27598.0534624214),
        ((0.1, 2, 3, 8000, 0, 1), -2416.91842900302),
        ((0.0, 2, 10, 1000), -100.0),
        ((0.0, 2, 10, 1000, 500), -150.0),
    ])
    def test_the_values(self, args, expected):
        assert close(PPMT(*args).value, expected)

    def test_the_spreadsheet_s_errors(self):
        assert PPMT(0.1, 4, 3, 8000).value == '#NUM!'
        assert PPMT(0.1, 0.5, 3, 8000).value == '#NUM!'

    def test_a_small_rate_loses_no_digits(self):
        # (1 + rate) ** nper - 1 by expm1, as the spreadsheet computes it:
        # at 1e-12 the subtraction alone keeps four digits, at 3e-16 none
        # (LibreOffice's values).
        assert close(PPMT(1e-9, 2, 10, 1000).value, -99.99999965)
        assert close(PPMT(1e-12, 2, 10, 1000).value, -99.99999999965)
        assert close(PPMT(1e-12, 2, 10, 1000, 0, 1).value, -99.99999999955)
        assert close(PPMT(3e-16, 2, 10, 1000).value, -99.9999999999999)
        assert close(PPMT(1e-9, 7.5, 10, 1000, 300).value, -130.00000026)
        assert close(CUMPRINC(1e-12, 360, 125000, 13, 24, 0).value, -4166.66666599167)

    @pytest.mark.parametrize("when", [0, 1])
    def test_interest_and_principal_make_the_payment_and_pay_the_loan_off(self, when):
        payment = PMT(0.1, 3, 8000, 0, when).value
        principal = [PPMT(0.1, p, 3, 8000, 0, when).value for p in (1, 2, 3)]
        interest = [IPMT(0.1, p, 3, 8000, 0, when).value for p in (1, 2, 3)]
        for i, p in zip(interest, principal):
            assert i + p == pytest.approx(payment)
        assert sum(principal) == pytest.approx(-8000.0)


class TestRate:
    @pytest.mark.parametrize("args, expected", [
        ((48, -200, 8000), 0.00770147248823326),
        ((48, -200, 8000, 0, 1), 0.00805298192393902),
        ((10, -100, 0, 1500), 0.0873205217799398),
        ((48, -200, 8000, 0, 0, 0.05), 0.00770147248820193),
        ((10.5, -100, 800), 0.0504359702865107),
        ((5, 50, -1000), -0.335280089136234),
    ])
    def test_the_values(self, args, expected):
        assert close(RATE(*args).value, expected)

    def test_the_rate_found_gives_the_payment_back(self):
        rate = RATE(360, -1000, 125000).value
        assert PMT(rate, 360, 125000).value == pytest.approx(-1000.0)
        rate = RATE(360, -1000, 125000, 5000, 1).value
        assert PMT(rate, 360, 125000, 5000, 1).value == pytest.approx(-1000.0)

    def test_a_loan_paid_back_exactly_has_a_rate_of_zero(self):
        # A rate too small to move 1 + rate is 0 (LibreOffice stops at
        # -7e-12), and the schedule reads it as no interest.
        assert RATE(10, -100, 1000).value == 0.0
        assert RATE(10, -100, 1000, guess=0.5).value == 0.0
        assert RATE(1000, -1, 1000).value == 0.0
        rate = RATE(10, -100, 1000).value
        assert IPMT(rate, 2, 10, 1000).value == 0.0 and PPMT(rate, 2, 10, 1000).value == -100.0

    def test_the_spreadsheet_s_iteration_and_its_further_starts(self):
        # A few payments the size of the loan: Newton from 10% lands below
        # -100%, and the spreadsheet starts again at 20%, 5%, 30%, ... as
        # the engine does. A thousand-period term settles inside the 150
        # steps. The numbers are LibreOffice's.
        assert close(RATE(5, -1000, 1000).value, 0.965948236645501)
        assert close(RATE(12, -1000, 1000).value, 0.999755500937318)
        assert close(RATE(5.5, -1000, 1000).value, 0.976413003900949)
        assert close(RATE(1200, -1, 1000).value, 0.000313468830444138)
        assert close(RATE(2, -1, 1000).value, -0.967873270798253)
        assert close(RATE(1, -2000, 1000).value, 1.0)
        # The rate settled on is refined to full precision: the payment
        # comes back to the last digits (LibreOffice stops within its step,
        # at 0.000999954352834482).
        rate = RATE(10000, -1, 1000).value
        assert close(rate, 0.000999954352834482)
        assert PMT(rate, 10000, 1000).value == pytest.approx(-1.0, rel=1e-12)
        # An annuity due's balance is x (400 x^2 - 600 x - 600) in x = 1 + r:
        # the rate is the quadratic's root, never the x = 0 at -100%.
        assert close(RATE(3, -600, 1000, 0, 1).value,
                     (600 + math.sqrt(600 ** 2 + 4 * 400 * 600)) / 800 - 1)

    def test_an_end_that_is_no_rate_is_a_start_that_found_none(self):
        # LibreOffice's iteration can stop at a point and show it as the
        # rate: -100%, where an annuity due's balance and a loan's with no
        # future value vanish trivially, or a point where the balance is
        # merely small. The engine takes neither: it tries the next start,
        # and with none left it is #NUM!.
        assert RATE(10, 0, 1000).value == '#NUM!'                  # LibreOffice: -98.2%
        assert RATE(36, 50, 0, 0, 1).value == '#NUM!'              # LibreOffice: -100.0%
        assert RATE(360, 0, -1000, 0, 1, 0.05).value == '#NUM!'    # LibreOffice: -10.4%
        assert RATE(1200, -10, 0).value == '#NUM!'                 # LibreOffice: 80%
        # From 10% the iteration ends at -100% (LibreOffice shows -100%);
        # a further start finds the rate the payments carry.
        for args in [(120, -100, 1000, 0, 1), (48, -200, 1000, 0, 1)]:
            rate = RATE(*args).value
            assert isinstance(rate, float) and rate > 0.1, args
            assert PV(rate, args[0], args[1], 0, 1).value == pytest.approx(1000.0)

    def test_a_given_guess_is_the_only_start(self):
        # LibreOffice starts again only from its own default; from a guess
        # given, an iteration that runs off is #NUM! (#N/A in LibreOffice).
        assert RATE(10, -100, 800, guess=-0.5).value == '#NUM!'
        assert RATE(10, -100, 1000, 0, 0, -2).value == '#NUM!'
        assert close(RATE(10, -100, 800, guess=0.05).value, RATE(10, -100, 800).value)

    def test_with_no_flow_at_all_every_rate_balances_and_the_guess_is_the_answer(self):
        assert RATE(10, 0, 0).value == 0.1                         # LibreOffice: 0.1
        assert RATE(10, 0, 0, guess=0.05).value == 0.05
        # One period repaid as it opens balances at any rate too
        # (LibreOffice's iteration wanders there on rounding, to 82.5%).
        assert RATE(1, -1000, 1000, 0, 1).value == 0.1

    def test_the_spreadsheet_s_errors(self):
        assert RATE(0, -100, 1000).value == '#NUM!'
        assert RATE(-5, -100, 1000).value == '#NUM!'
        assert RATE(10, 100, 1000).value == '#NUM!'               # every flow the same way
        # Payments five times the loan: no start settles (LibreOffice: #N/A).
        assert RATE(2, -5000, 1000).value == '#NUM!'
        assert RATE(12, -5000, 1000).value == '#NUM!'

    def test_no_unit_and_the_guess_written_only_when_given(self):
        n = mo.Variable(48, display_name="n")
        pmt = mo.Variable(-200.0, display_name="pmt", unit="$")
        loan = mo.Variable(8000.0, display_name="loan", unit="$")
        out = RATE(n, pmt, loan)
        assert out._unit is None and out.value_type == 'float'
        assert out.formula == "RATE(n, pmt, loan)" and len(out._expr.args) == 3
        assert RATE(n, pmt, loan, guess=0.05).formula == "RATE(n, pmt, loan, 0, 0, 0.05)"

    def test_a_series_nper_is_a_row_of_rates(self):
        out = RATE(mo.Variable([8, 9, 10]), -100, 800)
        assert out.var_type == 'list'
        assert out.value == [pytest.approx(RATE(n, -100, 800).value) for n in (8, 9, 10)]


class TestCumulative:
    @pytest.mark.parametrize("fn, args, expected", [
        (CUMIPMT, (MONTHLY, 360, 125000, 13, 24, 0), -11135.2321307508),
        (CUMIPMT, (MONTHLY, 360, 125000, 1, 1, 0), -937.5),
        (CUMIPMT, (MONTHLY, 360, 125000, 1, 12, 1), -10201.3328844946),
        (CUMIPMT, (MONTHLY, 360, 125000, 13.7, 24.2, 0), -11135.2321307508),
        (CUMPRINC, (MONTHLY, 360, 125000, 13, 24, 0), -934.107123420899),
        (CUMPRINC, (MONTHLY, 360, 125000, 1, 1, 0), -68.2782711809784),
        (CUMPRINC, (MONTHLY, 360, 125000, 1, 12, 1), -1778.16017175526),
    ])
    def test_the_values(self, fn, args, expected):
        assert close(fn(*args).value, expected)

    def test_the_sum_of_the_periods(self):
        interest = sum(IPMT(MONTHLY, p, 360, 125000).value for p in range(13, 25))
        principal = sum(PPMT(MONTHLY, p, 360, 125000).value for p in range(13, 25))
        assert CUMIPMT(MONTHLY, 360, 125000, 13, 24).value == pytest.approx(interest)
        assert CUMPRINC(MONTHLY, 360, 125000, 13, 24).value == pytest.approx(principal)
        whole = CUMIPMT(MONTHLY, 360, 125000, 1, 360).value + \
            CUMPRINC(MONTHLY, 360, 125000, 1, 360).value
        assert whole == pytest.approx(PMT(MONTHLY, 360, 125000).value * 360)

    def test_the_type_is_read_whole_and_written_always(self):
        assert CUMIPMT(MONTHLY, 360, 125000, 1, 12, 0.5).value == \
            CUMIPMT(MONTHLY, 360, 125000, 1, 12, 0).value
        rate = mo.Variable(MONTHLY, display_name="rate")
        out = CUMIPMT(rate, 360, 125000.0, 1, 12)
        assert out.formula == "CUMIPMT(rate, 360, 125000.0, 1, 12, 0)"
        assert len(out._expr.args) == 6
        for fn in (CUMIPMT, CUMPRINC):
            given_none = fn(rate, 360, 125000.0, 1, 12, when=None)
            assert given_none.formula == f"{fn.__name__}(rate, 360, 125000.0, 1, 12, 0)"
            assert given_none.value == fn(rate, 360, 125000.0, 1, 12).value

    def test_a_period_is_read_as_the_cell_shows_it(self):
        # 12.999999999999998 shows as 13, and the spreadsheet counts from 13.
        assert close(CUMIPMT(MONTHLY, 360, 125000, 12.999999999999998, 24, 0).value,
                     -11135.2321307508)
        assert close(CUMPRINC(MONTHLY, 360, 125000, 12.999999999999998, 24, 0).value,
                     -934.107123420899)

    @pytest.mark.parametrize("args", [
        (0.0, 360, 125000, 1, 12, 0),
        (-0.01, 360, 125000, 1, 12, 0),
        (MONTHLY, 0, 125000, 1, 12, 0),
        (MONTHLY, 360, 0, 1, 12, 0),
        (MONTHLY, 360, -125000, 1, 12, 0),
        (MONTHLY, 360, 125000, 0, 12, 0),
        (MONTHLY, 360, 125000, 5, 3, 0),
        (MONTHLY, 360, 125000, 1, 361, 0),
        (MONTHLY, 360, 125000, 1, 12, 2),
    ])
    def test_the_spreadsheet_s_errors(self, args):
        assert CUMIPMT(*args).value == '#NUM!'
        assert CUMPRINC(*args).value == '#NUM!'


# ══════════════════════════════════════════════════════════════════════
# Depreciation
# ══════════════════════════════════════════════════════════════════════


class TestSln:
    def test_the_values(self):
        assert SLN(30000, 7500, 10).value == 2250.0
        assert SLN(-30000, 7500, 10).value == -3750.0
        assert SLN(1000, 2000, 10).value == -100.0
        assert SLN(mo.Variable([30000.0, 20000.0]), 7500, 10).value == [2250.0, 1250.0]

    def test_a_life_of_zero_divides_by_zero(self):
        assert SLN(30000, 7500, 0).value == '#DIV/0!'


class TestDb:
    @pytest.mark.parametrize("period, expected", [
        (1, 186083.333333333), (2, 259639.416666667), (3, 176814.44275),
        (4, 120410.63551275), (5, 81999.6427841828), (6, 55841.7567360285),
        (7, 15845.0984738481),
    ])
    def test_the_documented_schedule_with_seven_months_in_the_first_year(self, period, expected):
        assert close(DB(1000000, 100000, 6, period, 7).value, expected)

    def test_twelve_months_when_month_is_omitted(self):
        assert close(DB(1000000, 100000, 6, 6).value, 46722.5182806209)
        assert DB(1000000, 100000, 6, 7).value == 0.0        # no months left for a stub
        assert DB(1000000, 100000, 6, 6, 12).value == DB(1000000, 100000, 6, 6).value

    def test_the_period_and_the_month_are_read_whole(self):
        assert DB(1000000, 100000, 6, 1.5, 7).value == DB(1000000, 100000, 6, 1, 7).value
        assert DB(1000000, 100000, 6, 2.5, 7).value == DB(1000000, 100000, 6, 2, 7).value
        assert DB(1000000, 100000, 6, 6.5, 7).value == DB(1000000, 100000, 6, 7, 7).value
        assert DB(1000000, 100000, 6, 2, 7.9).value == DB(1000000, 100000, 6, 2, 7).value

    def test_no_salvage_is_the_whole_cost_at_once(self):
        assert DB(1000, 0, 5, 1).value == 1000.0
        assert DB(1000, 0, 5, 2).value == 0.0

    @pytest.mark.parametrize("args, expected", [
        # A period that reads as 1 is the first period's charge, even past
        # a life shorter than that; the stub is for a later period.
        ((1000, 100, 0.5, 1), 990.0),
        ((1000, 100, 0.5, 1.5, 6), 495.0),
        ((1000, 100, 0.9, 1, 6), 461.5),
        ((1000, 100, 0.9, 1.9, 6), 461.5),
        ((1000, 100, 1, 1.5, 6), 450.0),
        ((1000, 100, 1.2, 1.5, 6), 426.5),
        ((1000, 0, 0.5, 1, 1), 83.3333333333333),
        ((1000, 100, 0.75, 1, 6), 477.0),
        ((1000, 100, 1, 2, 6), 247.5),
        ((1000, 100, 1.2, 2, 6), 244.59775),
        ((1000, 100, 1.5, 2.5, 6), 238.44375),
    ])
    def test_a_life_under_two_periods(self, args, expected):
        assert close(DB(*args).value, expected)

    def test_the_rate_period_and_month_are_read_as_the_cell_shows_them(self):
        # 1 - 318 / 2400 is 0.8675, held as 0.86749999...: the sheet shows
        # 0.8675 and rounds it to 0.868.
        assert close(DB(2400, 318, 1, 1).value, 2083.2)
        assert close(DB(2400, 318, 1, 2, 7).value, 428.502666666667)
        # A period or a month a hair under 3 shows as 3.
        assert close(DB(1000, 100, 5, 2.9999999999999996).value, 146.921409)
        assert close(DB(1000, 100, 5, 2.9999999999999996, 7).value, 182.72040525)
        assert close(DB(1000, 100, 5, 1, 2.9999999999999996).value, 92.25)

    @pytest.mark.parametrize("args", [
        (1000000, 100000, 6, 8, 7),
        (0, 0, 5, 1),
        (-1000, -100, 5, 1),
        (1000, -1, 5, 1),
        (1000, 2000, 5, 1),
        (1000, 100, 0, 1),
        (1000, 100, 5, 0),
        (1000, 100, 5, 0.5),            # Excel's domain; LibreOffice computes 0
        (1000, 100, 5, 1, 0),
        (1000, 100, 5, 1, 13),
    ])
    def test_the_spreadsheet_s_errors(self, args):
        assert DB(*args).value == '#NUM!'

    def test_a_series_period_is_a_schedule_row(self):
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=7)
        with m:
            m.cost = mo.Variable(1000000.0, display_name="cost", unit="$")
            m.dep = DB(m.cost, 100000.0, 6, mo.time.index, 7)
        assert m.dep.value == [pytest.approx(DB(1000000, 100000, 6, p, 7).value)
                               for p in range(1, 8)]
        assert str(m.dep._unit) == "$" and m.dep.var_type == 'list'


class TestDdb:
    @pytest.mark.parametrize("args, expected", [
        ((2400, 300, 3650, 1), 1.31506849315065),
        ((2400, 300, 120, 1), 40.0),
        ((2400, 300, 10, 1), 480.0),
        ((2400, 300, 10, 2, 1.5), 306.0),
        ((2400, 300, 10, 10), 22.1225472000002),
        ((2400, 300, 10, 1.5), 429.325051679959),
        ((1000, 100, 2, 1, 3), 900.0),
        ((1000, 100, 2, 2, 3), 0.0),
        ((0, 0, 5, 1), 0.0),
    ])
    def test_the_values(self, args, expected):
        assert close(DDB(*args).value, expected)

    def test_the_book_value_stops_at_the_salvage_value(self):
        charges = [DDB(2400, 300, 10, p).value for p in range(1, 11)]
        assert sum(charges) == pytest.approx(2400 - 300)
        assert charges[-1] == pytest.approx(22.1225472) and charges[0] == 480.0

    @pytest.mark.parametrize("args", [
        (2400, 300, 10, 11),
        (2400, 300, 10, 0.5),
        (1000, 0, 5, 0),
        (1000, 0, 5, 1, 0),
        (1000, 0, 5, 1, -2),
        (1000, 2000, 5, 1),
        (-1000, 0, 5, 1),
        (1000, -1, 5, 1),
        (1000, 0, 0, 1),
    ])
    def test_the_spreadsheet_s_errors(self, args):
        assert DDB(*args).value == '#NUM!'

    def test_a_series_period_is_a_schedule_row(self):
        row = DDB(2400.0, 300.0, 10, mo.Variable([1, 2, 3]))
        assert row.value == pytest.approx([480.0, 384.0, 307.2])
        assert row.var_type == 'list' and row.value_type == 'float'


class TestSyd:
    def test_the_values(self):
        assert close(SYD(30000, 7500, 10, 1).value, 4090.90909090909)
        assert close(SYD(30000, 7500, 10, 10).value, 409.090909090909)
        assert close(SYD(30000, 7500, 10, 2.5).value, 3477.27272727273)
        assert sum(SYD(30000, 7500, 10, p).value for p in range(1, 11)) == pytest.approx(22500)
        assert SYD(1000, 2000, 5, 1).value == pytest.approx(-333.333333333333)

    @pytest.mark.parametrize("args", [
        # Outside Excel's domain, where LibreOffice computes the formula
        # as written: 0 the period after the life, negative after that.
        (30000, 7500, 10, 11),
        (30000, 7500, 10, 12),
        (30000, 7500, 10, 10.5),
        (30000, 7500, 10, 0),
        (30000, 7500, 10, 0.5),
        (30000, 7500, 0, 1),
        (30000, 7500, -2, 1),
    ])
    def test_the_spreadsheet_s_errors(self, args):
        assert SYD(*args).value == '#NUM!'

    def test_a_series_period_is_a_schedule_row(self):
        row = SYD(30000.0, 7500.0, 10, mo.Variable([1, 2, 10]))
        assert row.value == [pytest.approx(4090.909090909091), pytest.approx(3681.818181818182),
                             pytest.approx(409.0909090909091)]


# ══════════════════════════════════════════════════════════════════════
# The shared law: one call per period
# ══════════════════════════════════════════════════════════════════════


class TestOneCallPerPeriod:
    def test_a_blank_is_zero(self):
        blank = mo.Variable([None])
        assert SLN(blank, 100, 10).value == [-10.0]
        assert DDB(2400, blank, 10, 1).value == [480.0]
        assert IPMT(blank, 2, 10, 1000).value == [0.0]
        assert NPER(0.1, -100, blank).value == [0.0]
        assert RATE(10, -100, 1000, blank).value == [pytest.approx(0.0, abs=1e-12)]

    def test_numeric_text_is_its_number_and_other_text_is_value(self):
        # As the spreadsheet reads text in an argument (LibreOffice's values).
        words = mo.Variable(["abc", "4"])
        assert SLN(words, 0, 1).value == ['#VALUE!', 4.0]
        assert IPMT(0.1, "x", 3, 8000).value == '#VALUE!'
        assert RATE(48, -200, "8000").value == RATE(48, -200, 8000).value
        assert IPMT(0.1, "2", 3, 8000).value == IPMT(0.1, 2, 3, 8000).value
        assert DB(1000, 100, 5, 1, "7").value == DB(1000, 100, 5, 1, 7).value
        assert CUMIPMT(0.01, 12, 1000, 1, 2, "0").value == CUMIPMT(0.01, 12, 1000, 1, 2, 0).value
        assert SLN(" 1,234.5 ", 0, 1).value == 1234.5
        assert SLN("12%", 0, 1).value == pytest.approx(0.12)
        assert SLN("(1,000)", 0, 1).value == -1000.0

    def test_a_boolean_is_one_or_zero(self):
        assert SLN(True, False, True).value == 1.0
        assert NPER(0.1, -100, 1000, 0, True).value == NPER(0.1, -100, 1000, 0, 1).value
        assert IPMT(0.1, True, 3, 8000).value == IPMT(0.1, 1, 3, 8000).value

    def test_a_numpy_number_is_a_number(self):
        assert SLN(np.int64(1000), 100, 5).value == 180.0
        assert RATE(np.int64(48), np.float64(-200.0), 8000).value == RATE(48, -200, 8000).value
        out = IPMT(0.1, mo.Variable(list(np.arange(1, 4))), 3, 8000)
        assert out.value == [pytest.approx(IPMT(0.1, p, 3, 8000).value) for p in (1, 2, 3)]
        # Read as a plain number, so a division by zero and an overflow are
        # the spreadsheet's errors, never numpy's inf.
        assert SLN(np.float64(1000.0), 100, 0).value == '#DIV/0!'
        assert IPMT(np.float64(0.1), 2, 10000, 1000).value == '#NUM!'
        assert type(SLN(np.float64(1000.0), 100, 5).value) is float

    def test_an_error_cell_is_the_period_s_result(self):
        e = mo.Variable([2.0, 2.0]) / mo.Variable([1.0, 0.0])
        assert e.value == [2.0, '#DIV/0!']
        assert SLN(e, 0, 1).value == [2.0, '#DIV/0!']
        assert IPMT(0.1, 1, 3, e).value == [pytest.approx(-0.2), '#DIV/0!']
        assert RATE(2, -1, e).value == [pytest.approx(0.0, abs=1e-12), '#DIV/0!']
        assert SYD(e, 0, 1, 1).value == [2.0, '#DIV/0!']

    def test_scalars_only_is_one_number(self):
        out = SLN(30000, 7500, 10)
        assert out.var_type == 'scalar' and out.value_type == 'float'
        assert out.formula == "SLN(30000, 7500, 10)"
        assert IPMT(mo.Variable(0.1), mo.Variable(1), mo.Variable(3), 8000.0).var_type == 'scalar'

    def test_scalars_and_a_one_value_series_broadcast(self):
        assert DDB(mo.Variable([2400.0]), 300, 10, [1, 2]).value == pytest.approx([480.0, 384.0])
        assert IPMT(mo.Variable([0.1]), [1, 2], 3, mo.Variable(8000.0)).value == \
            [pytest.approx(-800.0), pytest.approx(IPMT(0.1, 2, 3, 8000).value)]
        assert SLN(mo.Variable([1000.0, 2000.0]), 0, mo.Variable([10])).value == [100.0, 200.0]

    def test_series_of_different_lengths_are_refused_as_the_operators_refuse_them(self):
        with pytest.raises(ValueError, match="different lengths"):
            IPMT(0.1, mo.Variable([1, 2, 3]), 3, mo.Variable([8000.0, 9000.0]))

    def test_tracked_operands_lift_per_track(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"),
                     default_grain="month", default_start="2026-01", default_periods=2)
        with m:
            m.loan = mo.Variable(plan=[8000.0, 8000.0], actual=[9000.0, None])
            m.i = IPMT(0.1, 1, 3, m.loan)
            m.d = SLN(m.loan, 0.0, 10)
        assert dict(m.i._value.items()) == {"plan": [-800.0, -800.0], "actual": [-900.0, 0.0]}
        assert dict(m.d._value.items()) == {"plan": [800.0, 800.0], "actual": [900.0, 0.0]}
        assert m.i.var_type == 'list'

    def test_the_result_is_a_float_whatever_the_operands(self):
        assert IPMT(mo.Variable(1), 1, 3, mo.Variable(8000)).value_type == 'float'
        assert SLN(mo.Variable([30000, 20000]), 7500, 10).value_type == 'float'
        assert DDB(mo.Variable(2400), 300, 10, 1).value_type == 'float'
        assert NPER(mo.Variable(1), -100, 1000).value_type == 'float'

    def test_an_amount_keeps_the_unit_the_amounts_share(self):
        cost = mo.Variable(2400.0, unit="$")
        salvage = mo.Variable(300.0, unit="$")
        hours = mo.Variable(300.0, unit="hr")
        loan = mo.Variable(8000.0, unit="$")
        rate = mo.Variable(0.1)
        assert str(SLN(cost, salvage, 10)._unit) == "$"
        assert str(SLN(cost, 300.0, 10)._unit) == "$"
        # A cost and a salvage in two units are refused, as + refuses them.
        with pytest.raises(ValueError, match="mo.SLN takes its values in one unit"):
            SLN(cost, hours, 10)
        for fn in (DB, DDB, SYD):
            assert str(fn(cost, salvage, 10, 1)._unit) == "$", fn
        assert str(IPMT(rate, 1, 3, loan)._unit) == "$"
        assert str(PPMT(rate, 1, 3, loan)._unit) == "$"
        assert str(CUMIPMT(rate, 3, loan, 1, 2)._unit) == "$"
        assert str(CUMPRINC(rate, 3, loan, 1, 2)._unit) == "$"
        assert NPER(rate, -200.0, loan)._unit is None
        assert RATE(3, -200.0, loan)._unit is None

    def test_the_formula_names_the_operands_and_the_ast_is_a_call(self):
        rate = mo.Variable(0.1, display_name="rate")
        per = mo.Variable([1, 2, 3], display_name="per")
        term = mo.Variable(3, display_name="term")
        loan = mo.Variable(8000.0, display_name="loan")
        out = IPMT(rate, per, term, loan)
        assert out.formula == "IPMT(rate, per, term, loan)"
        assert isinstance(out._expr, FuncCall) and out._expr.func == "IPMT"
        assert all(isinstance(a, VarRef) for a in out._expr.args)
        assert out._expr.args[1].var is per
        lit = DDB(loan, 300.0, term, 2)
        assert lit.formula == "DDB(loan, 300.0, term, 2)"
        assert isinstance(lit._expr.args[1], Literal) and lit._expr.args[1].value == 300.0
        assert isinstance(lit._expr.args[3], Literal) and lit._expr.args[3].value == 2

    @pytest.mark.parametrize("name", sorted(NAMES))
    def test_an_axised_operand_is_refused(self, name):
        region = mo.Variable(["N", "S"], display_name="region")
        by_region = mo.Variable([1.0, 2.0], indexed_by=[region])
        arity = {"NPER": 3, "RATE": 3, "SLN": 3, "IPMT": 4, "PPMT": 4, "DB": 4, "DDB": 4,
                 "SYD": 4, "CUMIPMT": 5, "CUMPRINC": 5}[name]
        with pytest.raises(ValueError, match=name):
            NAMES[name](by_region, *([1] * (arity - 1)))


# ══════════════════════════════════════════════════════════════════════
# The registry and a loop container
# ══════════════════════════════════════════════════════════════════════


class TestTheRegistry:
    @pytest.mark.parametrize("name", sorted(NAMES))
    def test_each_is_element_wise_with_a_bare_classic_name(self, name):
        assert is_elementwise(name) and not is_range_taking(name)
        assert spec(name).passthrough and not spec(name).bare
        assert file_prefix(name) == "" and excel_name(name) == name
        assert admits_in_loop(name, 3) and admits_in_loop(name, 6)

    def test_rate_has_its_helper_now(self):
        assert not spec("RATE").bare and spec("RATE").shape == "elementwise"
        assert "RATE" in financial.__all__ and callable(financial.RATE)


class TestInsideALoop:
    def test_a_schedule_runs_period_by_period(self):
        # The opening balance reads the closing one through a lag, so the
        # container is a loop; NPER reads the opening balance — a row of
        # the loop — and says how many payments are left.
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
        with m:
            m.i = mo.MultiVariable("I")
            m.i.rate = mo.Variable(0.1, display_name="rate")
            m.i.loan = mo.Variable(1000.0, display_name="loan")
            m.i.payment = PMT(m.i.rate, 3, m.i.loan).set_display_name("payment")
            m.d = mo.MultiVariable("D", loop=True)
            with m.d as d:
                d.open = mo.lag(d.close, fill=m.i.loan)
                d.interest = d.open * m.i.rate
                d.principal = PPMT(m.i.rate, mo.time.index, 3, m.i.loan)
                d.close = d.open + d.principal
                d.left = NPER(m.i.rate, m.i.payment, d.open)
                d.rate = RATE(d.left, m.i.payment, d.open)
        # Interest on the opening balance is IPMT, which carries the
        # payment's sign: money out.
        assert [-v for v in d.interest.value] == \
            [pytest.approx(IPMT(0.1, p, 3, 1000).value) for p in (1, 2, 3)]
        assert d.close.value[-1] == pytest.approx(0.0, abs=1e-9)
        assert d.left.value == [pytest.approx(3.0), pytest.approx(2.0), pytest.approx(1.0)]
        # The rate recovered from the periods left and the balance is the input rate.
        assert d.rate.value == pytest.approx([0.1, 0.1, 0.1])
