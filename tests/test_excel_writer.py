# SPDX-License-Identifier: Apache-2.0
"""Tests for excel_writer — the explicit-root emission path.

Every test builds an explicit root MultiVariable and calls ``.to_excel()``
on it. Components attach via ``parent.child = ...`` only.
"""

import pytest
from openpyxl import load_workbook

import modeleon as mo


class TestToExcel:
    def test_writes_file(self, tmp_path):
        model = mo.MultiVariable("Writes")
        model.test = mo.MultiVariable("Test", excel_props={'tab': True})
        model.test.revenue = mo.Variable(1_000_000, display_name="Revenue")

        model.to_excel(tmp_path / "out.xlsx")

        path = tmp_path / "out.xlsx"
        assert path.exists()

    def test_empty_model_raises(self, tmp_path):
        model = mo.MultiVariable("Empty")
        with pytest.raises(ValueError, match="Nothing to emit"):
            model.to_excel(tmp_path / "out.xlsx")

    def test_creates_tab_per_sheet(self, tmp_path):
        model = mo.MultiVariable("Tabs")
        model.assumptions = mo.MultiVariable(
            "Assumptions", rate=mo.Variable(0.25), excel_props={'tab': True},
        )
        model.results = mo.MultiVariable(
            "Results", profit=mo.Variable(100_000), excel_props={'tab': True},
        )

        model.to_excel(tmp_path / "out.xlsx")

        path = tmp_path / "out.xlsx"
        wb = load_workbook(path)
        assert set(wb.sheetnames) == {"Assumptions", "Results"}

    def test_input_variable_writes_value(self, tmp_path):
        model = mo.MultiVariable("Inputs")
        model.test = mo.MultiVariable("Test", excel_props={'tab': True})
        model.test.revenue = mo.Variable(1_000_000, display_name="Revenue")

        model.to_excel(tmp_path / "out.xlsx")

        path = tmp_path / "out.xlsx"
        wb = load_workbook(path)
        ws = wb["Test"]

        values = [
            (cell.coordinate, cell.value)
            for row in ws.iter_rows()
            for cell in row
            if cell.value is not None
        ]
        labels = [v for _, v in values if isinstance(v, str)]
        numbers = [v for _, v in values if isinstance(v, (int, float))]

        assert "Revenue" in labels
        assert 1_000_000 in numbers

    def test_formula_variable_writes_excel_formula(self, tmp_path):
        model = mo.MultiVariable("Formulas")
        model.test = mo.MultiVariable("Test", excel_props={'tab': True})
        model.test.revenue = mo.Variable(1_000_000, display_name="Revenue")
        model.test.cogs_pct = mo.Variable(0.6, display_name="COGS %")
        model.test.cogs = model.test.revenue * model.test.cogs_pct

        model.to_excel(tmp_path / "out.xlsx")

        path = tmp_path / "out.xlsx"
        wb = load_workbook(path)
        ws = wb["Test"]

        formulas = []
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas.append(cell.value)

        assert len(formulas) >= 1
        for f in formulas:
            assert "revenue" not in f.lower() or "*" in f

    def test_unsafe_sheet_name_sanitized(self, tmp_path):
        model = mo.MultiVariable("Unsafe")
        model.unsafe = mo.MultiVariable(
            "Tab:With/Slash", x=mo.Variable(1), excel_props={'tab': True},
        )

        model.to_excel(tmp_path / "out.xlsx")

        path = tmp_path / "out.xlsx"
        wb = load_workbook(path)
        tab = wb.sheetnames[0]
        assert ":" not in tab
        assert "/" not in tab
