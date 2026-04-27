# SPDX-License-Identifier: Apache-2.0
"""Tests for financial functions: IRR, NPV, XIRR, PMT, FV, PV."""

import pytest
from openpyxl import load_workbook

import modeleon as mo
from modeleon import IRR, NPV


@pytest.fixture(autouse=True)
def clear_registry():

    yield

class TestNPV:
    def test_npv_positive_cashflows(self):
        # NPV at 10% of [-100, 50, 60, 70] should be positive
        result = NPV(0.10, [-100, 50, 60, 70])
        assert result > 0

    def test_npv_zero_rate(self):
        # NPV at 0% is just the sum
        result = NPV(0.0, [-100, 50, 50, 50])
        assert result == pytest.approx(50.0)


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
