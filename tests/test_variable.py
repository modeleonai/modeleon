# SPDX-License-Identifier: Apache-2.0
"""Tests for Variable - the core DSL primitive."""

import pytest

from modeleon import Variable


@pytest.fixture(autouse=True)
def clear_registry():

    yield

class TestVariableConstruction:
    def test_scalar_value(self):
        v = Variable(1000)
        assert v._value == 1000
        assert v.is_formula is False
        assert v.formula is None

    def test_list_value(self):
        v = Variable([100, 200, 300])
        assert v._value == [100, 200, 300]
        assert v.var_type == "list"

    def test_dict_sugar_becomes_keyed_list(self):
        v = Variable({"Bear": 800, "Base": 1000, "Bull": 1200})
        assert v._value == [800, 1000, 1200]
        assert v._keys == ["Bear", "Base", "Bull"]

    def test_display_name(self):
        v = Variable(1000, display_name="Revenue")
        assert v._display_name == "Revenue"

    def test_value_and_formula_mutually_exclusive(self):
        with pytest.raises(ValueError):
            Variable(value=1000, formula="a + b")


class TestArithmetic:
    def test_multiply_scalar(self):
        a = Variable(100)
        b = a * 0.6
        assert b._value == 60
        assert b.is_formula is True
        assert a.id in b.dependencies

    def test_multiply_two_variables(self):
        a = Variable(100)
        b = Variable(0.6)
        c = a * b
        assert c._value == 60
        assert a.id in c.dependencies
        assert b.id in c.dependencies

    def test_add(self):
        a = Variable(100)
        b = Variable(50)
        c = a + b
        assert c._value == 150

    def test_subtract(self):
        a = Variable(100)
        b = Variable(30)
        c = a - b
        assert c._value == 70

    def test_divide(self):
        a = Variable(100)
        b = Variable(4)
        c = a / b
        assert c._value == 25

    def test_chained_arithmetic(self):
        revenue = Variable(1000)
        cogs = revenue * 0.6
        profit = revenue - cogs
        assert profit._value == 400


class TestBroadcasting:
    def test_scalar_times_list(self):
        rate = Variable(0.25)
        values = Variable([100, 200, 300])
        result = values * rate
        assert result._value == [25.0, 50.0, 75.0]

    def test_list_plus_list_same_length(self):
        a = Variable([1, 2, 3])
        b = Variable([10, 20, 30])
        result = a + b
        assert result._value == [11, 22, 33]

    def test_scalar_plus_scalar(self):
        a = Variable(10)
        b = Variable(20)
        result = a + b
        assert result._value == 30


class TestSafeDivision:
    def test_divide_by_zero_returns_excel_error(self):
        a = Variable(100)
        b = Variable(0)
        result = a / b
        assert result._value == "#DIV/0!"


class TestCopy:
    """Variable.copy() and the __copy__ dunder produce equivalent results."""

    def test_copy_method(self):
        original = Variable(1000, display_name="Revenue")
        original._python_name = "revenue"
        duplicate = original.copy()
        assert duplicate._value == 1000
        assert duplicate.formula == "revenue.copy()"

    def test_copy_copy_dunder(self):
        import copy as _copy
        original = Variable([1, 2, 3])
        original._python_name = "nums"
        duplicate = _copy.copy(original)
        assert duplicate._value == [1, 2, 3]
        assert duplicate.formula == "nums.copy()"

    def test_both_paths_produce_equivalent_formulas(self):
        import copy as _copy
        v = Variable(100)
        v._python_name = "v"
        assert v.copy().formula == _copy.copy(v).formula


class TestPyFormula:
    """Variable(pyformula=...) wraps a value produced by arbitrary Python,
    captures its rendering on .pyformula, and marks the Variable as an
    input (is_formula=False). Used by downstream code-to-UI round trips."""

    def test_scalar_pyformula_sets_value(self):
        v = Variable(pyformula=42)
        assert v._value == 42
        assert v.pyformula == "42"
        assert v.is_formula is False

    def test_pyformula_with_list_value(self):
        v = Variable(pyformula=[1, 2, 3])
        assert v._value == [1, 2, 3]
        assert v.pyformula == "[1, 2, 3]"
        assert v.is_formula is False

    def test_pyformula_wrapping_variable_tracks_dependency(self):
        source = Variable(1000)
        source._python_name = "revenue"
        wrapped = Variable(pyformula=source)
        assert wrapped._value == 1000
        assert wrapped.pyformula == "revenue"
        assert wrapped.is_formula is False
        assert source in wrapped._dependency_refs

    def test_pyformula_captures_expression_result(self):
        """A Python expression result is accepted: .pyformula holds its repr."""
        def my_formula(a, b, c):
            return a * b + c

        a, b, c = Variable(10), Variable(2), Variable(5)
        result = my_formula(a._value, b._value, c._value)
        v = Variable(pyformula=result)
        assert v._value == 25
        assert v.pyformula == "25"

    def test_pyformula_mutually_exclusive_with_value_and_formula(self):
        with pytest.raises(ValueError):
            Variable(pyformula=1, value=2)
        with pytest.raises(ValueError):
            Variable(pyformula=1, formula="a + b")


class TestPythonName:
    """``python_name`` is set at attribute-assignment time
    (``mv.attr = var``) or via the explicit setter — never via frame
    introspection. Unattached Variables have ``python_name == None``."""

    def test_unattached_variable_has_no_python_name(self):
        v = Variable(1)
        assert v.python_name is None

    def test_attribute_assignment_sets_python_name(self):
        import modeleon as mo
        mv = mo.MultiVariable("M")
        mv.revenue = Variable(1_000_000)
        assert mv.revenue.python_name == 'revenue'

    def test_python_name_property_is_read_only(self):
        import pytest
        v = Variable(1)
        with pytest.raises(AttributeError):
            v.python_name = 'custom'


class TestDependenciesAsProperty:
    """dependencies is derived from _dependency_refs (the source of truth)."""

    def test_input_variable_has_no_dependencies(self):
        v = Variable(1000)
        assert v.dependencies == set()
        assert v._dependency_refs == []

    def test_arithmetic_populates_refs(self):
        a = Variable(10)
        b = Variable(20)
        c = a + b
        assert c._dependency_refs == [a, b]
        assert c.dependencies == {a.id, b.id}

    def test_dependencies_reflect_attached_names(self):
        """Adopted Variables get their python_name from the attribute
        slot. ``c.dependencies`` echoes the names the parent gave them."""
        import modeleon as mo
        mv = mo.MultiVariable("M")
        mv._python_name = "mv"      # crystallizes mv's path to ROOT.child("mv")
        mv.a = Variable(10)
        mv.b = Variable(20)
        mv.c = mv.a + mv.b
        # Adopted Variables resolve under the parent's path: ``mv.a``, ``mv.b``.
        assert mv.c.dependencies == {"mv.a", "mv.b"}

    def test_dependencies_is_read_only(self):
        v = Variable(1)
        with pytest.raises(AttributeError):
            v.dependencies = {"something"}

    def test_dependencies_mutating_refs_propagates(self):
        a = Variable(1)
        b = Variable(2)
        a._dependency_refs.append(b)
        assert a.dependencies == {b.id}
