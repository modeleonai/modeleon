# SPDX-License-Identifier: Apache-2.0
"""Unit propagation through arithmetic and Excel emission."""

import pytest
from openpyxl import load_workbook

import modeleon as mo


class TestBasicConstruction:
    def test_unit_stored_on_variable(self):
        v = mo.Variable(10, unit="$")
        assert str(v._unit) == "$"

    def test_compound_unit_parsed(self):
        v = mo.Variable(10, unit="$/hr")
        assert str(v._unit) == "$/hr"

    def test_no_unit_is_none(self):
        v = mo.Variable(10)
        assert v._unit is None


class TestMultiplication:
    def test_same_unit_composes(self):
        a = mo.Variable(5, unit="hr")
        b = mo.Variable(2, unit="hr")
        assert str((a * b)._unit) == "hr^2"

    def test_different_units_compose(self):
        price = mo.Variable(50, unit="$")
        hours = mo.Variable(10, unit="hr")
        assert str((price * hours)._unit) == "$ * hr"

    def test_scalar_preserves_unit(self):
        h = mo.Variable(10, unit="hr")
        assert str((h * 2)._unit) == "hr"
        assert str((2 * h)._unit) == "hr"


class TestDivision:
    def test_division_simplifies(self):
        revenue = mo.Variable(1000, unit="$")
        hours = mo.Variable(10, unit="hr")
        assert str((revenue / hours)._unit) == "$/hr"

    def test_cancellation_yields_dimensionless(self):
        a = mo.Variable(10, unit="$")
        b = mo.Variable(2, unit="$")
        assert (a / b)._unit is None

    def test_compound_division_cancels(self):
        cost = mo.Variable(100, unit="$")
        rate = mo.Variable(5, unit="$/hr")
        # $ / ($/hr) = hr
        assert str((cost / rate)._unit) == "hr"


class TestAddition:
    def test_same_unit_keeps_unit(self):
        a = mo.Variable(10, unit="$")
        b = mo.Variable(20, unit="$")
        assert str((a + b)._unit) == "$"

    def test_mixed_units_raise(self):
        a = mo.Variable(10, unit="$")
        b = mo.Variable(20, unit="hr")
        with pytest.raises(ValueError, match="incompatible units"):
            _ = a + b

    def test_subtract_mixed_raise(self):
        a = mo.Variable(10, unit="$")
        b = mo.Variable(20, unit="hr")
        with pytest.raises(ValueError, match="incompatible units"):
            _ = a - b

    def test_scalar_plus_unitful_keeps_unit(self):
        a = mo.Variable(10, unit="$")
        # 5 is dimensionless/scalar — treat as compatible, keep unit
        assert str((a + 5)._unit) == "$"


class TestPower:
    def test_integer_exponent_raises_unit(self):
        t = mo.Variable(3, unit="hr")
        assert str((t ** 2)._unit) == "hr^2"
        assert str((t ** 3)._unit) == "hr^3"

    def test_zero_exponent_dimensionless(self):
        t = mo.Variable(3, unit="hr")
        # hr^0 = dimensionless
        assert (t ** 0)._unit is None

    def test_float_exponent_drops_unit(self):
        t = mo.Variable(3, unit="hr")
        # Fractional/non-int exponents have no clean unit rendering
        assert (t ** 0.5)._unit is None


class TestNegation:
    def test_unary_minus_preserves_unit(self):
        a = mo.Variable(10, unit="$")
        assert str((-a)._unit) == "$"


class TestExcelEmission:
    def test_unit_appears_in_label(self, tmp_path):
        m = mo.MultiVariable("UT")
        m.u = mo.MultiVariable("U", excel_props={'tab': True})
        m.u.revenue = mo.Variable(1000, unit="$")
        m.to_excel(tmp_path / "u.xlsx")
        path = tmp_path / "u.xlsx"
        ws = load_workbook(path)["U"]
        labels = [
            cell.value for row in ws.iter_rows() for cell in row
            if isinstance(cell.value, str)
        ]
        assert any("($)" in label for label in labels)

    def test_no_unit_no_suffix(self, tmp_path):
        m = mo.MultiVariable("UT")
        m.u = mo.MultiVariable("U", excel_props={'tab': True})
        m.u.plain = mo.Variable(1000)
        m.to_excel(tmp_path / "u.xlsx")
        path = tmp_path / "u.xlsx"
        ws = load_workbook(path)["U"]
        labels = [
            cell.value for row in ws.iter_rows() for cell in row
            if isinstance(cell.value, str)
        ]
        assert not any("(" in label for label in labels)
