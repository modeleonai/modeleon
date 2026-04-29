# SPDX-License-Identifier: Apache-2.0
"""Tests for the day-1-coverage Excel functions — dates, text, loan math."""

from datetime import date

import pytest

import modeleon as mo


class TestDates:
    def test_year_scalar(self):
        assert mo.YEAR(mo.Variable(date(2025, 3, 14)))._value == 2025

    def test_month_from_iso_string(self):
        assert mo.MONTH(mo.Variable('2025-07-01'))._value == 7

    def test_day(self):
        assert mo.DAY(mo.Variable(date(2025, 12, 31)))._value == 31

    def test_edate_shifts_months(self):
        assert mo.EDATE(mo.Variable(date(2025, 1, 15)), 3)._value == date(2025, 4, 15)

    def test_edate_negative(self):
        assert mo.EDATE(mo.Variable(date(2025, 3, 15)), -2)._value == date(2025, 1, 15)

    def test_eomonth_same_month(self):
        assert mo.EOMONTH(mo.Variable(date(2025, 1, 15)))._value == date(2025, 1, 31)

    def test_eomonth_next(self):
        assert mo.EOMONTH(mo.Variable(date(2025, 1, 15)), 1)._value == date(2025, 2, 28)

    def test_today(self):
        result = mo.TODAY()
        assert result._value == date.today()
        assert result.formula == 'TODAY()'

    def test_year_list(self):
        dates = mo.Variable([date(2025, 1, 1), date(2026, 1, 1), date(2027, 1, 1)])
        assert mo.YEAR(dates)._value == [2025, 2026, 2027]


class TestText:
    def test_len(self):
        assert mo.LEN(mo.Variable('hello'))._value == 5

    def test_upper(self):
        assert mo.UPPER(mo.Variable('hello'))._value == 'HELLO'

    def test_lower(self):
        assert mo.LOWER(mo.Variable('HELLO'))._value == 'hello'

    def test_concat_variables(self):
        result = mo.CONCAT(mo.Variable('Q1 '), mo.Variable('Revenue'))
        assert result._value == 'Q1 Revenue'

    def test_concat_mixed(self):
        result = mo.CONCAT(mo.Variable('Total: '), 42)
        assert result._value == 'Total: 42'

    def test_concat_all_scalar_still_returns_variable(self):
        # mo.X(...) always returns a Variable — even for all-literal inputs,
        # the Excel cell holds =CONCAT("a", "b", "c"), not the collapsed text.
        result = mo.CONCAT('a', 'b', 'c')
        assert isinstance(result, mo.Variable)
        assert result._value == 'abc'
        assert result.formula.startswith('CONCAT(')


class TestFinancial:
    def test_pmt_matches_excel(self):
        # 30-year fixed-rate mortgage, $300k at 6% APR — Excel: -1798.6515...
        result = mo.PMT(0.06 / 12, 360, 300_000)
        assert result._value == pytest.approx(-1798.6515, rel=1e-4)

    def test_fv_annuity(self):
        # $500/month into 7%-return account for 30 years — Excel: 609985.50
        result = mo.FV(0.07 / 12, 360, -500)
        assert result._value == pytest.approx(609_985.50, rel=1e-4)

    def test_pv_annuity(self):
        # 20 years of $1000/yr at 5% — Excel: 12462.21...
        result = mo.PV(0.05, 20, -1000)
        assert result._value == pytest.approx(12_462.21, rel=1e-4)

    def test_pmt_zero_rate(self):
        # Zero-interest loan — payment is just principal / periods
        assert mo.PMT(0, 12, 1200)._value == pytest.approx(-100.0)

    def test_pmt_renders_formula(self):
        assert mo.PMT(0.05, 10, 1000).formula.startswith('PMT(')


class TestExcelOutput:
    """Sanity check that new functions render as real Excel formulas end-to-end."""

    def test_dates_and_text_round_trip(self, tmp_path):
        model = mo.MultiVariable('Demo')
        model.demo = mo.MultiVariable('Demo', excel_props={'tab': True})
        model.demo.launch = mo.Variable('2025-01-15', display_name='Launch')
        model.demo.y = mo.YEAR(model.demo.launch)
        model.demo.eom = mo.EOMONTH(model.demo.launch, 1)
        # Inputs to CONCAT/UPPER are attached to the model so the
        # renderer resolves them to real cells (no cross-scope fallback).
        model.demo.prefix = mo.Variable('q1 ')
        model.demo.suffix = mo.Variable('report')
        model.demo.header = mo.CONCAT(mo.UPPER(model.demo.prefix), model.demo.suffix)

        model.to_excel(str(tmp_path / 'demo.xlsx'))
        assert (tmp_path / 'demo.xlsx').exists()
