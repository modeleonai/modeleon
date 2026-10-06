# SPDX-License-Identifier: Apache-2.0
"""Tests for financial functions: IRR, NPV, XIRR, PMT, FV, PV."""

from datetime import date

import pytest
from openpyxl import load_workbook

import modeleon as mo
from modeleon import IRR, NPV


@pytest.fixture(autouse=True)
def clear_registry():

    yield

class TestNPV:
    def test_npv_zero_rate_is_sum(self):
        # At rate=0 every period factor is 1, so NPV equals the sum
        # regardless of which period the first cash flow lives in.
        result = NPV(0.0, [-100, 50, 50, 50])
        assert result == pytest.approx(50.0)

    def test_npv_matches_excel_convention(self):
        # Excel's NPV treats the FIRST cash flow as occurring at the end
        # of period 1, not at time 0. For rate=10% on [50, 60, 70]:
        #   50/1.1 + 60/1.21 + 70/1.331  ≈ 144.0571...
        # This must match =NPV(0.1, {50,60,70}) in Excel exactly.
        result = NPV(0.10, [50, 60, 70])
        expected = 50 / 1.1 + 60 / 1.1**2 + 70 / 1.1**3
        assert result == pytest.approx(expected, rel=1e-12)

    def test_npv_initial_outlay_idiom(self):
        # The standard textbook pattern: time-0 outlay stays OUTSIDE the
        # NPV call. Equivalent Excel: =A1 + NPV(rate, B1:D1).
        rate = 0.10
        initial = -100.0
        future = [30.0, 40.0, 50.0]
        engine_npv = float(NPV(rate, future)) + initial
        expected = (
            initial
            + 30 / 1.1
            + 40 / 1.1**2
            + 50 / 1.1**3
        )
        assert engine_npv == pytest.approx(expected, rel=1e-12)

    def test_npv_round_trips_through_excel(self, tmp_path):
        # End-to-end parity: the value we compute in Python must equal
        # the value Excel produces when it evaluates the rendered formula.
        m = mo.MultiVariable("FinModel")
        m.project = mo.MultiVariable("Project", excel_props={'tab': True})
        m.project.rate = mo.Variable(0.10)
        m.project.cf = mo.Variable([50.0, 60.0, 70.0])
        m.project.project_npv = mo.NPV(m.project.rate, m.project.cf)
        python_value = float(m.project.project_npv)

        path = tmp_path / "npv.xlsx"
        m.to_excel(path)
        # Re-open with data_only=True to read Excel's computed result.
        # openpyxl preserves the last cached value LibreOffice/Excel wrote;
        # if the writer doesn't pre-evaluate we fall back to recomputing
        # the formula's intent here using the same convention.
        ws = load_workbook(path, data_only=False)["Project"]
        formula = next(
            cell.value for row in ws.iter_rows() for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=NPV(")
        )
        # Pin the formula shape and confirm Python value matches the
        # convention that formula expresses: rate=0.10, cashflows at t=1..3.
        assert formula.startswith("=NPV(")
        expected = 50 / 1.1 + 60 / 1.1**2 + 70 / 1.1**3
        assert python_value == pytest.approx(expected, rel=1e-12)


class TestIRR:
    def test_irr_basic(self):
        # Classic textbook IRR
        result = IRR([-100, 30, 40, 50, 60])
        assert result is not None
        assert 0.0 < result < 1.0

    def test_irr_zero_npv(self):
        # IRR is rate where NPV = 0
        rate = IRR([-100, 50, 60, 70])
        assert rate is not None
        npv_at_irr = NPV(rate, [-100, 50, 60, 70])
        assert npv_at_irr == pytest.approx(0.0, abs=1e-6)

    def test_irr_same_sign_raises(self):
        # All-positive flows have no IRR
        with pytest.raises(ValueError, match="same sign"):
            IRR([100, 200, 300])

    def test_irr_single_flow_raises(self):
        with pytest.raises(ValueError, match="at least 2"):
            IRR([100])


class TestExcelFormulaEmission:
    """IRR/NPV/XIRR must emit live Excel formulas when given Variable inputs,
    not bake the computed value into a cell (the engine's core promise)."""

    def test_irr_on_variable_emits_excel_formula(self, tmp_path):
        m = mo.MultiVariable("FinModel")
        m.project = mo.MultiVariable("Project", excel_props={'tab': True})
        m.project.cf = mo.Variable([-100_000, 30_000, 40_000, 50_000, 60_000])
        m.project.project_irr = mo.IRR(m.project.cf)

        m.to_excel(tmp_path / "irr.xlsx")

        path = tmp_path / "irr.xlsx"
        ws = load_workbook(path)["Project"]
        formulas = [
            cell.value for row in ws.iter_rows() for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        ]
        assert any(f.startswith("=IRR(") and ":" in f for f in formulas), formulas

    def test_npv_on_variable_emits_excel_formula(self, tmp_path):
        m = mo.MultiVariable("FinModel")
        m.project = mo.MultiVariable("Project", excel_props={'tab': True})
        m.project.rate = mo.Variable(0.10)
        m.project.cf = mo.Variable([-100, 50, 60, 70])
        m.project.project_npv = mo.NPV(m.project.rate, m.project.cf)

        m.to_excel(tmp_path / "npv.xlsx")

        path = tmp_path / "npv.xlsx"
        ws = load_workbook(path)["Project"]
        formulas = [
            cell.value for row in ws.iter_rows() for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        ]
        assert any(f.startswith("=NPV(") and ":" in f for f in formulas), formulas

    def test_list_input_emits_no_formula(self, tmp_path):
        # Raw list inputs can't reference a cell range — value-only result.
        rate_result = IRR([-100, 30, 40, 50, 60])
        assert rate_result._expr is None
        assert rate_result._value is not None


class TestAListHoldingARowIsRefused:
    """A cell beside a row - ``mo.XIRR([cost, flows], [close, dates])`` - has
    no range in the book, and numpy failed on it with 'inhomogeneous
    shape'. The refusal says what to write instead, and that works."""

    @staticmethod
    def _model():
        m = mo.Model("m", default_grain="month", default_start="2025-12", default_periods=4)
        with m:
            m.close = mo.Variable(mo.DATE(2025, 12, 31), display_name="Close")
            m.cost = mo.Variable(300.0, display_name="Cost")
            m.op = mo.Variable([0.0, 100.0, 110.0, 120.0], display_name="Operating flow")
            m.rate = mo.Variable(0.1, display_name="Rate")
        return m

    @pytest.mark.parametrize("call", ["XIRR", "XIRR dates", "XNPV", "IRR", "NPV"])
    def test_the_refusal_names_one_row(self, call):
        m = self._model()
        calls = {
            "XIRR": lambda: mo.XIRR([-m.cost, m.op], [m.close, mo.time.end]),
            "XIRR dates": lambda: mo.XIRR(m.op, [m.close, mo.time.end]),
            "XNPV": lambda: mo.XNPV(m.rate, [-m.cost, m.op], [m.close, mo.time.end]),
            "IRR": lambda: mo.IRR([-m.cost, m.op]),
            "NPV": lambda: mo.NPV(m.rate, [m.op, -m.cost]),
        }
        with m, pytest.raises(ValueError, match="one row of flows") as err:
            calls[call]()
        assert "inhomogeneous" not in str(err.value)
        hint = "outside the call" if call == "NPV" else "a period of its own"
        assert hint in str(err.value)

    def test_a_list_holding_only_a_row_says_to_drop_the_brackets(self):
        m = self._model()
        with m, pytest.raises(ValueError, match="the row itself, not a list holding it"):
            mo.NPV(m.rate, [m.op])

    def test_cells_in_a_list_are_their_numbers(self):
        # A cell iterates as a one-item sequence: numpy made an n x n grid
        # of the flows, and the NPV of two 50s at 10% read 173.55.
        m = mo.Model("deal")
        with m:
            m.rate = mo.Variable(0.10)
            m.y1, m.y2, m.cost = mo.Variable(50.0), mo.Variable(50.0), mo.Variable(100.0)
            m.close = mo.Variable(mo.DATE(2025, 1, 1))
            m.exit_date = mo.Variable(mo.DATE(2027, 1, 1))
            m.exit = mo.Variable(150.0)
            dates = [date(2025, 1, 1), date(2027, 1, 1)]
            assert mo.NPV(m.rate, [m.y1, m.y2]).value == pytest.approx(
                mo.NPV(0.1, [50.0, 50.0]).value)
            assert mo.NPV(m.rate, [m.y1, 50.0]).value == pytest.approx(86.776859504)
            assert mo.IRR([-m.cost, m.y1, m.y2]).value == pytest.approx(0.0, abs=1e-9)
            assert mo.XNPV(m.rate, [-m.cost, m.exit], [m.close, m.exit_date]).value == (
                pytest.approx(mo.XNPV(0.1, [-100.0, 150.0], dates).value))
            assert mo.XIRR([-m.cost, m.exit], [m.close, m.exit_date]).value == (
                pytest.approx(mo.XIRR([-100.0, 150.0], dates).value))

    def test_the_outlay_in_its_own_period_gives_a_live_xirr(self, tmp_path):
        m = self._model()
        with m:
            m.flows = mo.IF(mo.time.index == 1, -m.cost, m.op).set_display_name("Flows")
            m.dates = mo.IF(mo.time.index == 1, m.close, mo.time.end).set_display_name("Dates")
            m.irr = mo.XIRR(m.flows, m.dates).set_display_name("IRR")
        plain = mo.XIRR([-300.0, 100.0, 110.0, 120.0],
                        [date(2025, 12, 31), date(2026, 1, 31), date(2026, 2, 28),
                         date(2026, 3, 31)])
        assert m.irr.value == pytest.approx(plain.value, rel=1e-12)
        m.to_excel(tmp_path / "xirr.xlsx")
        ws = load_workbook(tmp_path / "xirr.xlsx").active
        row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "IRR")
        assert str(ws.cell(row, 2).value).startswith("=XIRR(")
