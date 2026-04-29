# SPDX-License-Identifier: Apache-2.0
"""Tests for ``_repr_html_`` — Jupyter-rendered HTML for Variable / MV / Sheet / Model."""

import pytest

import modeleon as mo


class TestVariableRepr:
    def test_scalar_renders_table(self):
        v = mo.Variable(1_000_000, display_name="Revenue", unit="$")
        html = v._repr_html_()
        assert "<table" in html
        assert "Revenue ($)" in html
        assert "1,000,000" in html

    def test_list_renders_one_row_per_period(self):
        v = mo.Variable([200, 250, 300, 400], display_name="New Customers")
        html = v._repr_html_()
        assert "New Customers" in html
        for n in (200, 250, 300, 400):
            assert f"{n}" in html

    def test_unit_appears_in_header(self):
        v = mo.Variable(99, display_name="ARPU", unit="$")
        assert "ARPU ($)" in v._repr_html_()

    def test_no_unit_no_parens(self):
        v = mo.Variable(0.04, display_name="Monthly Churn")
        html = v._repr_html_()
        assert "Monthly Churn" in html
        assert "Monthly Churn (" not in html

    def test_float_formatting(self):
        # large numbers get thousands separator and no decimals
        big = mo.Variable(1234567.89)
        assert "1,234,568" in big._repr_html_()
        # small fractions get more precision
        small = mo.Variable(0.001234)
        assert "0.0012" in small._repr_html_()


class TestMultiVariableRepr:
    def test_sheet_renders_grid(self):
        sheet = mo.MultiVariable("P&L", excel_props={'tab': True})
        sheet.revenue = mo.Variable(1_000_000, display_name="Revenue")
        sheet.cogs = mo.Variable(600_000, display_name="COGS")
        html = sheet._repr_html_()
        assert "P&L" in html or "P&amp;L" in html or "P&L" in html
        assert "Revenue" in html
        assert "COGS" in html
        assert "1,000,000" in html

    def test_section_header_for_nested_mv(self):
        class Cohort(mo.MultiVariableClass):
            def compute(self, start=100):
                self.users = mo.Variable([start, start * 0.9, start * 0.81])

        sheet = mo.MultiVariable("Cohorts", excel_props={'tab': True})
        sheet.jan = Cohort(start=100)
        sheet.feb = Cohort(start=200)

        html = sheet._repr_html_()
        # Each cohort should appear as a section header
        assert "Jan" in html
        assert "Feb" in html


class TestModelRepr:
    def test_renders_each_sheet(self):
        m = mo.MultiVariable("Forecast")
        m.assumptions = mo.MultiVariable("Assumptions", excel_props={'tab': True})
        m.assumptions.arpu = mo.Variable(99, display_name="ARPU", unit="$")
        m.plan = mo.MultiVariable("Plan", excel_props={'tab': True})
        m.plan.revenue = mo.Variable(1_000_000, display_name="Revenue")
        html = m._repr_html_()
        assert "Forecast" in html
        assert "Assumptions" in html
        assert "Plan" in html
        assert "ARPU ($)" in html
        assert "Revenue" in html

    def test_empty_model_renders_gracefully(self):
        m = mo.MultiVariable("Empty")
        html = m._repr_html_()
        assert "Empty" in html
        assert "empty" in html


class TestVariableReprAddresses:
    """A lone Variable's repr lays out via the same LayoutEngine the
    rest of the renderer uses, so the cell carries a real Excel
    address keyed off the Variable's name. Pin these so the
    `_make_variable_sheet` pseudo-sheet wrapper doesn't drift back to
    a hard-coded `Variable!A1`/no-address shape."""

    def test_anonymous_variable_default_tab_name(self):
        # No display_name, no python_name → fall back to "Variable".
        v = mo.Variable(42)
        html = v._repr_html_()
        assert 'data-cell="Variable!B1"' in html

    def test_named_variable_uses_display_name_as_tab(self):
        v = mo.Variable(100, display_name="Revenue")
        html = v._repr_html_()
        assert 'data-cell="Revenue!B1"' in html

    def test_adopted_variable_uses_humanized_python_name(self):
        m = mo.Model("m")
        m.cogs = mo.Variable(60)
        html = m.cogs._repr_html_()
        assert 'data-cell="Cogs!B1"' in html

    def test_recurrence_chain_gets_per_period_addresses(self):
        r = mo.recurrence(start=1000, formula="{prev}*1.05", periods=4)
        html = r._repr_html_()
        for col in ("B", "C", "D", "E"):
            assert f'data-cell="Variable!{col}1"' in html


class TestModelHtmlPositional:
    """``model_html`` treats the rendered root as the workbook itself:
    first-depth sub-MVs become tabs, deeper sub-MVs become sections in
    the panel's flat layout. The test here pins the user's `t.h = acme`
    scenario — rendering ``t`` exposes one `Acme` tab where the inner
    sheet-marked `pnl` falls through as a section at row 1."""

    def test_nested_sheet_renders_as_section_in_parent_tab(self):
        acme = mo.Model("acme")
        with mo.MultiVariable(excel_props={"tab": True}) as pnl:
            pnl.revenue = mo.Variable(1_000_000)
            pnl.cogs = pnl.revenue * mo.Variable(0.6)
        acme.pnl = pnl
        t = mo.Model("t")
        t.h = acme

        html = t._repr_html_()

        # The `Pnl` section header lands on row 1 of the Acme tab —
        # the user's documented expectation. Cell addresses are
        # zero-indexed against this flat layout, so cogs's formula
        # references revenue via ``B2`` (row 2 in the panel).
        assert 'data-cell="Acme!B2"' in html
        assert 'data-cell="Acme!B3"' in html
        # cogs formula references revenue's row in the flat layout.
        assert "=B2 * 0.6" in html

    def test_cross_tab_formula_uses_qualified_address(self):
        # A formula on the Outputs tab that references Variables on the
        # Inputs tab must render with sheet-qualified addresses
        # (``=Inputs!B1 * Inputs!B2``), not fall through to inlined
        # literals — this matches what ``to_excel`` produces. The fix
        # is a workbook-wide translator built from the union of every
        # panel's per-panel addresses.
        m = mo.Model("m")
        m.inputs = mo.MultiVariable("Inputs", excel_props={"tab": True})
        m.inputs.price = mo.Variable(100.0, display_name="Price")
        m.inputs.volume = mo.Variable(50, display_name="Volume")
        m.outputs = mo.MultiVariable("Outputs", excel_props={"tab": True})
        m.outputs.revenue = (m.inputs.price * m.inputs.volume).set_display_name(
            "Revenue"
        )

        html = m._repr_html_()

        assert "=Inputs!B1 * Inputs!B2" in html
        # Sanity: the literal-fallback shape must NOT be present —
        # otherwise the cross-tab path silently regressed.
        assert "=100.0 * 50" not in html

    def test_each_panel_carries_its_name_for_stacked_mode(self):
        # The "Stacked" toolbar toggle hides the tab bar and reveals
        # every panel at once — so each panel's title must stay in
        # the markup as a ``.mo-panel-title`` div (CSS hides it in
        # tabbed mode, shows it in stacked mode). Previously the panel
        # titles were suppressed entirely, leaving stacked mode
        # identifying nothing.
        m = mo.Model("m")
        m.inputs = mo.MultiVariable("Inputs", excel_props={"tab": True})
        m.inputs.price = mo.Variable(100.0, display_name="Price")
        m.outputs = mo.MultiVariable("Outputs", excel_props={"tab": True})
        m.outputs.revenue = m.inputs.price * mo.Variable(2)

        html = m._repr_html_()

        import re

        titles = re.findall(
            r'<div class="mo-panel-title"[^>]*>([^<]+)</div>', html,
        )
        assert titles == ["Inputs", "Outputs"]
