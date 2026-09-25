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

    def test_date_construct(self):
        assert mo.DATE(2025, 4, 15)._value == date(2025, 4, 15)

    def test_date_month_overflow_rolls_year(self):
        # the common DATE(YEAR, MONTH+3, DAY) quarter-step past December
        assert mo.DATE(2025, 15, 10)._value == date(2026, 3, 10)

    def test_date_from_variables(self):
        d = mo.Variable(date(2025, 10, 20))
        assert mo.DATE(mo.YEAR(d), mo.MONTH(d), mo.DAY(d))._value == date(2025, 10, 20)

    def test_date_list_elementwise(self):
        years = mo.Variable([2025, 2026])
        assert mo.DATE(years, 1, 1)._value == [date(2025, 1, 1), date(2026, 1, 1)]

    def test_days360_european_quarter(self):
        assert mo.DAYS360(date(2025, 1, 1), date(2025, 4, 1), True)._value == 90

    def test_days360_european_full_year(self):
        assert mo.DAYS360(date(2025, 1, 1), date(2026, 1, 1), True)._value == 360

    def test_days360_european_day31_becomes_30(self):
        # 31 Jan -> 30 Apr on the European basis: both day-31s clamp to 30
        assert mo.DAYS360(date(2025, 1, 31), date(2025, 3, 31), True)._value == 60

    def test_days360_us_default_method(self):
        assert mo.DAYS360(date(2025, 1, 31), date(2025, 2, 28))._value == 28

    def test_days360_us_february_month_end_start_counts_as_the_30th(self):
        # US/NASD as Excel computes it: a start on the last day of
        # February becomes the 30th, so a day-31 end then clamps too.
        assert mo.DAYS360(date(2023, 2, 28), date(2023, 3, 31))._value == 30
        assert mo.DAYS360(date(2024, 2, 29), date(2024, 3, 31))._value == 30
        assert mo.DAYS360(date(2023, 2, 28), date(2023, 3, 28))._value == 28

    def test_days360_us_february_28_of_a_leap_year_is_not_month_end(self):
        assert mo.DAYS360(date(2024, 2, 28), date(2024, 3, 31))._value == 33

    def test_days360_european_ignores_february_month_end(self):
        assert mo.DAYS360(date(2023, 2, 28), date(2023, 3, 31), True)._value == 32

    def test_days360_list_elementwise(self):
        starts = mo.Variable([date(2025, 1, 1), date(2025, 4, 1)])
        ends = mo.Variable([date(2025, 4, 1), date(2025, 7, 1)])
        assert mo.DAYS360(starts, ends, True)._value == [90, 90]


class TestLogical:
    def test_and_scalar(self):
        assert mo.AND(mo.Variable(5) > 3, mo.Variable(2) < 10)._value is True
        assert mo.AND(mo.Variable(1) > 3, mo.Variable(2) < 10)._value is False

    def test_or_scalar(self):
        assert mo.OR(mo.Variable(1) > 3, mo.Variable(2) < 10)._value is True

    def test_not_scalar(self):
        assert mo.NOT(mo.Variable(5) > 3)._value is False

    def test_and_elementwise(self):
        x = mo.Variable([1, 5, 9, 12])
        assert mo.AND(x > 3, x < 11)._value == [False, True, True, False]

    def test_choose_scalar(self):
        assert mo.CHOOSE(2, 'base', 'bull', 'bear')._value == 'bull'

    def test_choose_out_of_range(self):
        assert mo.CHOOSE(9, 'a', 'b')._value == '#VALUE!'

    def test_choose_elementwise_phase_switch(self):
        flag = mo.Variable([0, 0, 1, 1])
        assert mo.CHOOSE(flag + 1, 'Ф', 'П')._value == ['Ф', 'Ф', 'П', 'П']

    def test_choose_scalar_index_over_series_choices_is_a_series(self):
        # The scenario switch: one scenario number picks a whole series.
        base = mo.Variable([100, 110, 120])
        bull = mo.Variable([150, 170, 190])
        result = mo.CHOOSE(mo.Variable(2), base, bull)
        assert result._value == [150, 170, 190]
        assert result.var_type == 'list'
        assert mo.CHOOSE(1, base, bull)._value == [100, 110, 120]

    def test_choose_scalar_choice_repeats_beside_series_choices(self):
        # The result is a series whichever choice the index picks, so the
        # row keeps its shape when the scenario number changes.
        bull = mo.Variable([150, 170, 190])
        assert mo.CHOOSE(1, 125, bull)._value == [125, 125, 125]

    def test_choose_out_of_range_over_series_choices(self):
        base = mo.Variable([100, 110, 120])
        bull = mo.Variable([150, 170, 190])
        assert mo.CHOOSE(3, base, bull)._value == ['#VALUE!'] * 3

    def test_choose_one_value_series_repeats_beside_longer_series(self):
        # A one-value series is written as one cell that every period reads.
        one = mo.Variable([7])
        bull = mo.Variable([150, 170, 190])
        assert mo.CHOOSE(1, one, bull)._value == [7, 7, 7]

    @pytest.mark.parametrize('short', [[1, 2], []])
    def test_choose_series_choices_of_different_lengths_are_refused(self, short):
        # The workbook has no cell for a short series' missing periods, so
        # any value Python made up for them would disagree with Excel.
        base = mo.Variable(short)
        bull = mo.Variable([150, 170, 190, 210])
        with pytest.raises(ValueError, match='different lengths'):
            mo.CHOOSE(1, base, bull)

    def test_logical_render_to_excel(self, tmp_path):
        import openpyxl
        model = mo.Model('Logic')
        model.x = mo.Variable([1, 5, 9], display_name='X')
        model.ok = mo.AND(model.x > 3, model.x < 11)
        model.ok._display_name = 'Ok'
        out = tmp_path / 'logic.xlsx'
        model.to_excel(str(out))
        ws = openpyxl.load_workbook(str(out)).active
        formulas = [c.value for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith('=')]
        assert any(f.startswith('=AND(') for f in formulas), formulas


class TestText:
    def test_len(self):
        assert mo.LEN(mo.Variable('hello'))._value == 5

    def test_concat_elementwise_period_label(self):
        quarter = mo.Variable([1, 2, 3])
        year = mo.Variable([2022, 2022, 2022])
        phase = mo.Variable(['Ф', 'Ф', 'П'])
        label = mo.CONCAT(quarter, 'Q ', year, ' ', phase)
        assert label._value == ['1Q 2022 Ф', '2Q 2022 Ф', '3Q 2022 П']

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

    def test_concat_renders_as_ampersand_not_concat_function(self, tmp_path):
        """CONCAT must emit the ``&`` operator, never the ``CONCAT()``
        function — CONCAT is an Excel-2016 future function that needs an
        ``_xlfn.`` prefix and otherwise shows as ``@CONCAT`` / ``#NAME?``."""
        import openpyxl

        model = mo.Model('Label')
        model.q = mo.Variable([1, 2], display_name='Q')
        model.y = mo.Variable([2022, 2022], display_name='Y')
        model.label = mo.CONCAT(model.q, 'Q ', model.y)
        model.label._display_name = 'Label'
        out = tmp_path / 'label.xlsx'
        model.to_excel(str(out))

        ws = openpyxl.load_workbook(str(out)).active
        formulas = [c.value for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith('=')]
        label_formulas = [f for f in formulas if '&' in f]
        assert label_formulas, formulas
        assert not any('CONCAT' in f for f in formulas), formulas

    def test_date_literal_in_formula_renders_as_date_function(self, tmp_path):
        """A date literal embedded in a formula must emit ``=DATE(y, m, d)``
        — a bare ``2025-01-01`` would be read by Excel as ``2025 - 1 - 1``.
        The recurrence start at period 0 is such an embedded literal."""
        import openpyxl

        model = mo.Model('Dates')
        model.date = mo.recurrence(date(2025, 1, 1), 'EDATE(prev, 3)', periods=3,
                                   display_name='Date')
        out = tmp_path / 'dates.xlsx'
        model.to_excel(str(out))

        ws = openpyxl.load_workbook(str(out)).active
        formulas = [c.value for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith('=')]
        assert any('DATE(2025, 1, 1)' in f for f in formulas), formulas

    def test_days360_schedule_renders_back_references(self, tmp_path):
        """DAYS360 over a shifted date axis emits previous-cell references,
        and the shift fill value keeps its date type (=DATE(...), not text)."""
        import openpyxl

        model = mo.Model('Sched')
        model.date = mo.recurrence(date(2025, 1, 1), 'EDATE(prev, 3)', periods=4,
                                   display_name='Date')
        model.days = mo.DAYS360(model.date.shift(1, fill_value=date(2025, 1, 1)),
                                model.date, True)
        model.days._display_name = 'Days'
        out = tmp_path / 'sched.xlsx'
        model.to_excel(str(out))

        ws = openpyxl.load_workbook(str(out)).active
        formulas = [c.value for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith('=')]
        # day-count formulas reference prior period cells, never textual junk
        days_cells = [f for f in formulas if 'DAYS360' in f]
        assert days_cells
        assert all('"' not in f for f in days_cells), days_cells  # no quoted-text dates
        assert any('DATE(2025, 1, 1)' in f for f in days_cells)   # fill kept its type


def _soffice() -> str | None:
    """Path to the LibreOffice command line, or None when not installed."""
    import shutil
    from pathlib import Path

    found = shutil.which('soffice') or shutil.which('libreoffice')
    if found:
        return found
    mac = Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
    return str(mac) if mac.exists() else None


def _recalculated_rows(xlsx, workdir) -> dict[str, list[str]]:
    """Open the workbook in LibreOffice, let it calculate every formula,
    and return the first sheet as ``{label in column A: other cells}``."""
    import csv
    import subprocess

    soffice = _soffice()
    if soffice is None:
        pytest.skip('LibreOffice is not installed')
    profile = (workdir / 'lo-profile').as_uri()  # private profile: parallel runs don't clash
    subprocess.run(
        [soffice, f'-env:UserInstallation={profile}', '--headless',
         '--convert-to', 'csv', '--outdir', str(workdir), str(xlsx)],
        check=True, capture_output=True, timeout=120,
    )
    with open(workdir / (xlsx.stem + '.csv'), newline='', encoding='utf-8') as fh:
        return {row[0]: row[1:] for row in csv.reader(fh) if row}


class TestChooseScenarioSwitch:
    """``mo.CHOOSE(scenario, base, bull)`` — one scenario number, series choices.

    Every period gets its own ``=CHOOSE(...)`` that points at the same
    scenario cell and at that period's cell of each choice, so changing
    the scenario number in the workbook switches the whole row.
    """

    @staticmethod
    def _model(scenario: int) -> mo.Model:
        m = mo.Model('Scenarios')
        m.scenario = mo.Variable(scenario, display_name='Scenario')
        m.base = mo.Variable([100, 110, 120], display_name='Base')
        m.bull = mo.Variable([150, 170, 190], display_name='Bull')
        m.revenue = mo.CHOOSE(m.scenario, m.base, m.bull)
        m.revenue._display_name = 'Revenue'
        m.floor = mo.CHOOSE(m.scenario, 125, m.bull)
        m.floor._display_name = 'Floor'
        return m

    def test_one_formula_per_period(self, tmp_path):
        import openpyxl

        out = tmp_path / 'scenarios.xlsx'
        self._model(2).to_excel(str(out))
        ws = openpyxl.load_workbook(str(out)).active
        rows = {row[0].value: row for row in ws.iter_rows()}
        idx = rows['Scenario'][1].coordinate
        base_row = rows['Base'][0].row
        bull_row = rows['Bull'][0].row
        got = [c.value for c in rows['Revenue'][1:4]]
        assert got == [
            f'=CHOOSE({idx}, {col}{base_row}, {col}{bull_row})' for col in 'BCD'
        ]
        assert [c.value for c in rows['Floor'][1:4]] == [
            f'=CHOOSE({idx}, 125, {col}{bull_row})' for col in 'BCD'
        ]

    @pytest.mark.parametrize('scenario', [1, 2])
    def test_workbook_recalculates_to_the_python_values(self, tmp_path, scenario):
        m = self._model(scenario)
        out = tmp_path / 'scenarios.xlsx'
        m.to_excel(str(out))
        calc = _recalculated_rows(out, tmp_path)
        for label, var in (('Revenue', m.revenue), ('Floor', m.floor)):
            excel = [float(v) if v else None for v in calc[label][:3]]
            assert excel == var._value, label


class TestISBLANK:
    """``ISBLANK`` — the only honest way to write an OPTIONAL value.

    A leaving date that has not happened yet must read as an EMPTY cell.
    Filling it with the end of the window says "dismissed on 31
    December", which an accountant will act on. But comparing against an
    empty operand answers all-False, so "no upper bound" cannot be said
    without a blank test — and Excel's own answer to that is ISBLANK.
    """

    def test_empty_is_blank_and_a_value_is_not(self) -> None:
        assert mo.ISBLANK(mo.Variable(None)).value is True
        assert mo.ISBLANK(mo.Variable("")).value is True
        assert mo.ISBLANK(mo.Variable(0)).value is False
        assert mo.ISBLANK(mo.Variable("Иванов")).value is False

    def test_zero_is_not_blank(self) -> None:
        """The distinction the whole thing rests on: a rate of 0 % is a
        stated value, an unset rate is not."""
        assert mo.ISBLANK(mo.Variable(0)).value is False
        assert mo.ISBLANK(mo.Variable(0.0)).value is False

    def test_element_wise_over_a_list(self) -> None:
        got = mo.ISBLANK(mo.Variable([1, None, "", 3])).value
        assert got == [False, True, True, False]

    def test_renders_the_excel_function(self) -> None:
        assert mo.ISBLANK(mo.Variable(None)).formula.startswith("ISBLANK(")

    def test_an_open_ended_window_counts_to_the_end(self) -> None:
        """Typical use: an employee with no leaving date
        stays in staff through the window; one with a date stops."""
        start = mo.Variable([mo.DATE(2026, m, 1) for m in range(1, 13)])
        end = mo.Variable(mo.EOMONTH(start))
        hired = mo.DATE(2026, 4, 10)

        def months(left: mo.Variable) -> int:
            on = mo.IF(
                end >= hired,
                mo.IF(mo.OR(mo.ISBLANK(left), start <= left), 1, 0),
                0,
            )
            return sum(on.value)

        # Hired 10 April, still employed → April through December.
        assert months(mo.Variable(None)) == 9
        # Hired 10 April, left 15 August → April through August.
        assert months(mo.DATE(2026, 8, 15)) == 5
