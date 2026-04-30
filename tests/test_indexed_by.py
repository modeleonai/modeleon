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

    def test_copy_shares_axis_identity(self):
        # A copy of a Variable indexed by an axis still references the
        # *same* axis object — not a clone of it. Two-step broadcasting
        # works because axis identity is preserved.
        months = mo.Variable(['Jan'])
        v = mo.Variable([1], indexed_by=[months])
        c = v.copy()
        assert c._indexed_by[0] is months


class TestArithmeticPreservesAxes:
    """Every arithmetic operator must propagate ``_indexed_by`` —
    not just ``*`` and ``+``. Pin each operator individually so a
    regression in any single op gets caught."""

    def setup_method(self) -> None:
        self.months = mo.Variable(['Jan', 'Feb', 'Mar'])
        self.a = mo.Variable([10.0, 20.0, 30.0], indexed_by=[self.months])
        self.b = mo.Variable([2.0, 2.0, 2.0], indexed_by=[self.months])

    def test_add(self):
        assert (self.a + self.b)._indexed_by == (self.months,)

    def test_sub(self):
        assert (self.a - self.b)._indexed_by == (self.months,)

    def test_mul(self):
        assert (self.a * self.b)._indexed_by == (self.months,)

    def test_div(self):
        assert (self.a / self.b)._indexed_by == (self.months,)

    def test_pow(self):
        assert (self.a ** 2)._indexed_by == (self.months,)

    def test_floordiv(self):
        assert (self.a // self.b)._indexed_by == (self.months,)

    def test_mod(self):
        assert (self.a % self.b)._indexed_by == (self.months,)


class TestReflectedOpsPreserveAxes:
    """Reflected operators (``__rmul__``, ``__rsub__``, etc.) fire
    when the LEFT operand is a non-Variable. The result must still
    pick up the Variable's axes."""

    def test_scalar_minus_indexed(self):
        months = mo.Variable(['Jan', 'Feb'])
        v = mo.Variable([10, 20], indexed_by=[months])
        # ``__rsub__`` path
        result = 100 - v
        assert result._indexed_by == (months,)

    def test_scalar_div_indexed(self):
        months = mo.Variable(['Jan', 'Feb'])
        v = mo.Variable([10, 20], indexed_by=[months])
        result = 1 / v
        assert result._indexed_by == (months,)

    def test_scalar_times_indexed(self):
        months = mo.Variable(['Jan', 'Feb'])
        v = mo.Variable([1, 2], indexed_by=[months])
        result = 0.5 * v
        assert result._indexed_by == (months,)


class TestComparisonsPreserveAxes:
    """Comparison ops also build new Variables and must carry shape
    forward. The resulting Variable is the truthiness mask used by
    ``mo.where``-style conditionals downstream."""

    def test_lt(self):
        months = mo.Variable(['Jan', 'Feb'])
        a = mo.Variable([1, 2], indexed_by=[months])
        b = mo.Variable([2, 1], indexed_by=[months])
        assert (a < b)._indexed_by == (months,)

    def test_eq(self):
        months = mo.Variable(['Jan', 'Feb'])
        a = mo.Variable([1, 2], indexed_by=[months])
        assert (a == 1)._indexed_by == (months,)

    def test_ge_disjoint_axes(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        a = mo.Variable([1, 2], indexed_by=[months])
        b = mo.Variable([1, 2], indexed_by=[scenarios])
        # Comparison broadcasts the axis union just like arithmetic.
        assert (a >= b)._indexed_by == (months, scenarios)


class TestMultiStepBroadcast:
    """A chain of three ops over three different axes ends up with
    all three axes in the result's ``_indexed_by``. This is the
    cross-product case that any FP&A model with multiple dimensions
    will hit (months × scenarios × products)."""

    def test_three_disjoint_axes(self):
        months    = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        products  = mo.Variable(['A', 'B'])

        per_month    = mo.Variable([1.0, 2.0],  indexed_by=[months])
        per_scenario = mo.Variable([0.9, 1.1],  indexed_by=[scenarios])
        per_product  = mo.Variable([10.0, 20.0], indexed_by=[products])

        result = per_month * per_scenario * per_product
        assert result._indexed_by == (months, scenarios, products)

    def test_repeated_axis_in_chain_dedupes(self):
        months = mo.Variable(['Jan', 'Feb'])
        a = mo.Variable([1, 2], indexed_by=[months])
        b = mo.Variable([3, 4], indexed_by=[months])
        c = mo.Variable([5, 6], indexed_by=[months])
        # Same axis used at three sites should appear once in the
        # final ``_indexed_by``.
        result = (a + b) * c
        assert result._indexed_by == (months,)


class TestAdoptedAxisIdentity:
    """An axis Variable adopted into an MV keeps the same Python
    identity. Variables ``indexed_by`` it before or after adoption
    point at the same axis object; shape readouts agree."""

    def test_indexed_by_before_adoption(self):
        months = mo.Variable(['Jan', 'Feb'])
        revenue = mo.Variable([100, 110], indexed_by=[months])
        # Now adopt months into a Model.
        m = mo.Model('m')
        m.months = months
        # The axis Variable that revenue references is the same
        # object as ``m.months`` — adoption doesn't clone it.
        assert revenue._indexed_by[0] is m.months

    def test_indexed_by_after_adoption(self):
        m = mo.Model('m')
        m.months = mo.Variable(['Jan', 'Feb'])
        m.revenue = mo.Variable([100, 110], indexed_by=[m.months])
        assert m.revenue._indexed_by == (m.months,)


class TestExcelEmissionWithIndexedBy:
    """``to_excel`` must keep working for Variables that declare
    ``indexed_by``. Shape lives on the dep graph; layout still walks
    the MV tree and writes one row per Variable. No interaction
    needed for now — but the smoke test guards future regressions."""

    def test_indexed_by_variable_emits_without_error(self, tmp_path):
        m = mo.Model('m')
        m.months = mo.Variable(['Jan', 'Feb', 'Mar'])
        m.revenue = mo.Variable(
            [100.0, 110.0, 121.0],
            indexed_by=[m.months],
            display_name='Revenue',
        )

        m.to_excel(tmp_path / 'out.xlsx')
        assert (tmp_path / 'out.xlsx').exists()

    def test_arithmetic_intermediate_shape_survives_emission(self, tmp_path):
        m = mo.Model('m')
        m.months = mo.Variable(['Jan', 'Feb'])
        m.revenue = mo.Variable([100.0, 110.0], indexed_by=[m.months])
        m.cogs    = (m.revenue * 0.6).set_display_name('Cogs')
        # Shape inferred via broadcast survives until emission.
        assert m.cogs._indexed_by == (m.months,)
        m.to_excel(tmp_path / 'out.xlsx')


class TestEdgeCases:
    """Boundary cases we want pinned so regressions show up early."""

    def test_empty_axis_list(self):
        months = mo.Variable([])
        v = mo.Variable([], indexed_by=[months])
        assert v._indexed_by == (months,)
        assert v.shape.length == 0

    def test_indexed_by_kwarg_accepts_tuple(self):
        # Sequence acceptance — list, tuple, generator should all work.
        months = mo.Variable(['Jan'])
        v = mo.Variable([1], indexed_by=(months,))
        assert v._indexed_by == (months,)

    def test_shape_repr_uses_display_name(self):
        m = mo.Model('m')
        m.scenarios = mo.Variable(['Bear', 'Bull'])
        v = mo.Variable([1, 2], indexed_by=[m.scenarios])
        # ``scenarios`` was adopted under ``m`` — its python_name is set.
        assert 'scenarios' in repr(v.shape)

    def test_shape_length_is_product_of_axis_sizes(self):
        a = mo.Variable(['x', 'y', 'z'])
        b = mo.Variable(['p', 'q'])
        v = mo.Variable([[0, 0], [0, 0], [0, 0]], indexed_by=[a, b])
        assert v.shape.length == 6   # 3 × 2
