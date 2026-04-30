# SPDX-License-Identifier: Apache-2.0
"""Tests for ``indexed_by``-driven shape inference.

Replaces the old ``Shape / AxisProvider / MutableAxisProvider /
TimeBehavior`` machinery. Shape now lives in the dep graph: each
Variable declares which axes it's laid out along via
``indexed_by=[...]``; operator-built Variables get the union of
operand axes automatically. ``shape_of(var)`` is the read accessor.

These tests pin the contract of the new model: the read path
(``shape_of`` / ``var.shape``), the broadcast propagation through
ops, and the independence of ``_keys`` (keyed-list labels) from
``_indexed_by`` (axis declarations).
"""

from __future__ import annotations

import modeleon as mo
from modeleon.core.shape import Shape, shape_of


class TestDefaultShape:
    """Variables with no ``indexed_by`` declaration are scalar — the
    most common case for plain inputs and operator results that don't
    touch any axis."""

    def test_scalar_value_has_empty_axes(self):
        v = mo.Variable(0.25)
        assert shape_of(v) == Shape(())
        assert v.shape.is_scalar

    def test_list_value_without_indexed_by_is_scalar(self):
        # Old behavior would synthesize a positional axis here. New
        # behavior: lists alone don't imply an axis; the user has to
        # declare ``indexed_by=`` to opt in.
        v = mo.Variable([1, 2, 3])
        assert v.shape == Shape(())

    def test_formula_variable_default_scalar(self):
        v = mo.Variable(formula='x + 1')
        assert v.shape == Shape(())


class TestExplicitIndexedBy:
    """Pure inputs declare their axes via ``indexed_by=[…]``. The
    declaration is preserved on the Variable and read back via
    ``shape``."""

    def test_indexed_by_single_axis(self):
        months = mo.Variable(['Jan', 'Feb', 'Mar'])
        revenue = mo.Variable([100, 110, 121], indexed_by=[months])
        assert revenue._indexed_by == (months,)
        assert revenue.shape == Shape((months,))

    def test_indexed_by_two_axes(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        cube = mo.Variable(
            [[80, 100], [85, 110]],
            indexed_by=[months, scenarios],
        )
        assert cube._indexed_by == (months, scenarios)

    def test_indexed_by_default_empty_tuple(self):
        # Constructor default is no axes → tuple().
        v = mo.Variable(42)
        assert v._indexed_by == ()


class TestOperatorBroadcast:
    """Binary ops broadcast operand axes into the result's
    ``_indexed_by``. Same axis is deduped; non-overlapping axes
    compose; scalar × axis-bearing returns axis-bearing."""

    def test_scalar_times_axis_inherits_axis(self):
        months = mo.Variable(['Jan', 'Feb', 'Mar'])
        revenue = mo.Variable([100, 110, 121], indexed_by=[months])
        cogs = revenue * 0.6
        assert cogs._indexed_by == (months,)
        # Value broadcast still works.
        assert cogs._value == [60.0, 66.0, 72.6]

    def test_same_axis_dedupes(self):
        months = mo.Variable(['Jan', 'Feb'])
        a = mo.Variable([1, 2], indexed_by=[months])
        b = mo.Variable([3, 4], indexed_by=[months])
        result = a + b
        assert result._indexed_by == (months,)

    def test_disjoint_axes_compose(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        per_month = mo.Variable([100, 110], indexed_by=[months])
        per_scenario = mo.Variable([0.9, 1.1], indexed_by=[scenarios])
        # Broadcasting two disjoint axes yields the union — order is
        # left then right.
        combined = per_month * per_scenario
        assert combined._indexed_by == (months, scenarios)

    def test_unary_negation_preserves_indexed_by(self):
        months = mo.Variable(['Jan', 'Feb'])
        revenue = mo.Variable([100, 110], indexed_by=[months])
        loss = -revenue
        assert loss._indexed_by == (months,)


class TestShapeOf:
    """``shape_of(var)`` is the public accessor; equivalent to
    ``var.shape`` but importable as a function for use sites that
    don't have a Variable to dot-into."""

    def test_shape_of_matches_property(self):
        months = mo.Variable(['Jan'])
        v = mo.Variable([1], indexed_by=[months])
        assert shape_of(v) == v.shape

    def test_shape_of_scalar(self):
        v = mo.Variable(7)
        assert shape_of(v) == Shape(())

    def test_shape_of_2d(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        cube = mo.Variable(
            [[1, 2], [3, 4]], indexed_by=[months, scenarios],
        )
        assert shape_of(cube) == Shape((months, scenarios))


class TestKeysIndependent:
    """``_keys`` (keyed-list labels) and ``_indexed_by`` (axis
    declarations) are independent. Both can coexist on one Variable;
    each handles its own concern."""

    def test_keyed_variable_without_indexed_by(self):
        # Existing keyed-list behavior continues unchanged.
        v = mo.Variable({'US': 100, 'EU': 80, 'APAC': 60})
        assert v._keys == ['US', 'EU', 'APAC']
        assert v._value == [100, 80, 60]
        # No indexed_by declared.
        assert v._indexed_by == ()

    def test_keyed_variable_with_indexed_by(self):
        regions = mo.Variable(['US', 'EU', 'APAC'])
        v = mo.Variable(
            [100, 80, 60], keys=['US', 'EU', 'APAC'], indexed_by=[regions],
        )
        # Both fields populated, independent of each other.
        assert v._keys == ['US', 'EU', 'APAC']
        assert v._indexed_by == (regions,)


class TestCopyPreservesIndexedBy:
    """``Variable.copy()`` is the supported clone path. It must
    propagate ``_indexed_by`` so the copy reports the same shape."""

    def test_copy_preserves_axes(self):
        months = mo.Variable(['Jan', 'Feb', 'Mar'])
        v = mo.Variable([1, 2, 3], indexed_by=[months])
        c = v.copy()
        assert c._indexed_by == v._indexed_by
        assert c.shape == v.shape
