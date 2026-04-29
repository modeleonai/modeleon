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


class TestPositionalSheetSectionRule:
    """`to_excel` matches the HTML repr's positional rule: the root is
    the workbook (never its own tab), first-depth sub-MVs become tabs,
    deeper sub-MVs render as sections inside their containing tab —
    regardless of an inner MV's ``_is_sheet`` flag."""

    def test_root_is_not_a_tab_even_if_marked(self, tmp_path):
        # Even though `t` is marked `tab=True`, it's the workbook
        # itself; the only tab should come from its first-depth child.
        acme = mo.Model("acme")
        acme.revenue = mo.Variable(1_000_000)
        t = mo.Model("t", excel_props={"tab": True})
        t.h = acme

        t.to_excel(tmp_path / "t.xlsx")

        wb = load_workbook(tmp_path / "t.xlsx")
        assert wb.sheetnames == ["Acme"]

    def test_nested_sheet_marker_renders_as_section(self, tmp_path):
        # `pnl._is_sheet=True`, but it's depth 2 from the root passed
        # to `to_excel`, so it renders as a section in `Acme`, not a
        # separate tab. Section header lands at row 1; the variable
        # rows + cogs formula reference the flat layout.
        acme = mo.Model("acme")
        with mo.MultiVariable(excel_props={"tab": True}) as pnl:
            pnl.revenue = mo.Variable(1_000_000, display_name="Revenue")
            pnl.cogs = (pnl.revenue * mo.Variable(0.6)).set_display_name("Cogs")
        acme.pnl = pnl
        t = mo.Model("t")
        t.h = acme

        t.to_excel(tmp_path / "t.xlsx")

        wb = load_workbook(tmp_path / "t.xlsx")
        assert wb.sheetnames == ["Acme"]
        ws = wb["Acme"]
        cells = {
            cell.coordinate: cell.value
            for row in ws.iter_rows()
            for cell in row
            if cell.value is not None
        }
        assert cells == {
            "A1": "Pnl",  # nested sheet flattens to a section header
            "A2": "Revenue", "B2": 1_000_000,
            "A3": "Cogs",    "B3": "=B2 * 0.6",
        }

    def test_orphan_variables_get_synthesized_overview_tab(self, tmp_path):
        # When the root has no sub-MVs, only Variables, those orphans
        # land in a synthesized overview tab named after the root.
        pnl = mo.MultiVariable("PnL")
        pnl.revenue = mo.Variable(100, display_name="Revenue")
        pnl.cogs = (pnl.revenue * mo.Variable(0.6)).set_display_name("Cogs")

        pnl.to_excel(tmp_path / "p.xlsx")

        wb = load_workbook(tmp_path / "p.xlsx")
        assert wb.sheetnames == ["PnL"]


class TestVariableToExcel:
    """A lone ``Variable`` exports just like its HTML repr — one row
    on a tab named after the Variable. The pseudo-sheet wrapper used
    for layout is the same one ``variable_html`` uses, so file output
    and notebook output agree on tab name + cell shape."""

    def test_named_variable_writes_one_row(self, tmp_path):
        v = mo.Variable(100, display_name="Revenue")

        v.to_excel(tmp_path / "v.xlsx")

        wb = load_workbook(tmp_path / "v.xlsx")
        assert wb.sheetnames == ["Revenue"]
        ws = wb["Revenue"]
        cells = {
            cell.coordinate: cell.value
            for row in ws.iter_rows()
            for cell in row
            if cell.value is not None
        }
        assert cells == {"A1": "Revenue", "B1": 100}

    def test_recurrence_writes_period_chain(self, tmp_path):
        r = mo.recurrence(
            start=1000, formula="{prev}*1.05", periods=4, display_name="Users",
        )

        r.to_excel(tmp_path / "r.xlsx")

        wb = load_workbook(tmp_path / "r.xlsx")
        assert wb.sheetnames == ["Users"]
        ws = wb["Users"]
        formulas = {
            cell.coordinate: cell.value
            for row in ws.iter_rows()
            for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        }
        # Period 0 = literal start; later periods chain via {prev}.
        assert formulas == {
            "B1": "=1000",
            "C1": "=B1*1.05",
            "D1": "=C1*1.05",
            "E1": "=D1*1.05",
        }
