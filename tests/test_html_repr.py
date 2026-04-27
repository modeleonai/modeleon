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
