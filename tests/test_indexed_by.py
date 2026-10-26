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

import operator

import pytest

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
        # Flat row-major — the ONLY accepted cube spelling:
        # nested lists are ambiguous about which level is which axis.
        cube = mo.Variable(
            [80, 100, 85, 110],
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
        # left then right — AND the true cross-product VALUE (row-major,
        # last axis fastest), never a 2-cell diagonal under a 4-cell
        # label.
        combined = per_month * per_scenario
        assert combined._indexed_by == (months, scenarios)
        assert combined._value == [
            100 * 0.9, 100 * 1.1,   # Jan × (Bear, Bull)
            110 * 0.9, 110 * 1.1,   # Feb × (Bear, Bull)
        ]

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
            [1, 2, 3, 4], indexed_by=[months, scenarios],
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
        # 8 cells, row-major over (months, scenarios, products).
        want = [
            m * s * pr
            for m in (1.0, 2.0)
            for s in (0.9, 1.1)
            for pr in (10.0, 20.0)
        ]
        assert result._value == pytest.approx(want)

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
        v = mo.Variable([0, 0, 0, 0, 0, 0], indexed_by=[a, b])
        assert v.shape.length == 6   # 3 × 2


class TestAxisedConstructionInvariant:
    """A declared shape must be honored by the value."""

    def test_length_mismatch_rejected_with_teaching_error(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        with pytest.raises(ValueError, match="require 4 value"):
            mo.Variable([1.0, 2.0], indexed_by=[months, scenarios])

    def test_nested_lists_rejected(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        with pytest.raises(ValueError, match="FLAT row-major"):
            mo.Variable([[1.0, 2.0], [3.0, 4.0]],
                        indexed_by=[months, scenarios])

    def test_matching_flat_value_accepted(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        v = mo.Variable([1.0, 2.0, 3.0, 4.0], indexed_by=[months, scenarios])
        assert v._value == [1.0, 2.0, 3.0, 4.0]
        assert v._indexed_by == (months, scenarios)


class TestIdentityAlignedBroadcast:
    """The value under the union label is the cross-product."""

    def test_scalar_still_broadcasts(self):
        months = mo.Variable(['Jan', 'Feb'])
        v = mo.Variable([10.0, 20.0], indexed_by=[months])
        r = v * 2
        assert r._value == [20.0, 40.0]
        assert r._indexed_by == (months,)

    def test_same_axis_stays_elementwise(self):
        months = mo.Variable(['Jan', 'Feb'])
        a = mo.Variable([1.0, 2.0], indexed_by=[months])
        b = mo.Variable([10.0, 20.0], indexed_by=[months])
        assert (a + b)._value == [11.0, 22.0]

    def test_transposed_operands_align_by_axis_identity(self):
        # c is (months, scenarios); d is (scenarios, months). Same set,
        # different declared order — the values must align by identity,
        # not by position, or c + d is silently scrambled.
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        c = mo.Variable([1.0, 2.0, 3.0, 4.0], indexed_by=[months, scenarios])
        d = mo.Variable([10.0, 30.0, 20.0, 40.0], indexed_by=[scenarios, months])
        # d in (months, scenarios) order is [10, 20, 30, 40].
        total = c + d
        assert total._indexed_by == (months, scenarios)
        assert total._value == [11.0, 22.0, 33.0, 44.0]

    def test_axised_times_bigger_axised(self):
        months = mo.Variable(['Jan', 'Feb', 'Mar'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        m = mo.Variable([1.0, 2.0, 3.0], indexed_by=[months])
        s = mo.Variable([10.0, 100.0], indexed_by=[scenarios])
        r = m * s
        assert r._indexed_by == (months, scenarios)
        assert r._value == [10.0, 100.0, 20.0, 200.0, 30.0, 300.0]

    def test_comparison_crosses_axes_too(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        m = mo.Variable([1.0, 3.0], indexed_by=[months])
        s = mo.Variable([2.0, 2.0], indexed_by=[scenarios])
        r = m > s
        assert r._indexed_by == (months, scenarios)
        assert r._value == [False, False, True, True]


class TestRank1Guards:
    """Rank-1 guard — a single finite axis is not a time series.

    Without the guard, grain projection summed coordinates into quarters
    with no error; lag shifted across coordinates and dropped
    indexed_by. Every time-coupled operation refuses axised operands."""

    def _axised(self, display_name='Scenarios'):
        scenarios = mo.Variable(['Bear', 'Bull'])
        return mo.Variable([10.0, 20.0], indexed_by=[scenarios],
                           display_name=display_name)

    def test_time_is_none_for_axised(self):
        # In a windowed model an axised variable must NOT inherit the
        # ambient (start, grain) — its cells are coordinates, not months.
        m = mo.Model('m', default_grain='month', default_start='2025-01',
                     default_periods=2)
        with m:
            m.x = self._axised()
        assert m.x.time is None

    def test_at_grain_refuses(self):
        with pytest.raises(ValueError, match="finite axis"):
            self._axised().at('quarter')

    def test_projection_refuses_rank1(self):
        from modeleon.core.projection import project_variable
        with pytest.raises(ValueError, match="axised"):
            project_variable(self._axised(), 'quarter', {})

    def test_lag_refuses(self):
        from modeleon.functions.lag import lag
        with pytest.raises(ValueError, match="finite axis"):
            lag(self._axised())

    def test_cumsum_refuses(self):
        with pytest.raises(ValueError, match="finite axis"):
            mo.cumsum(self._axised())

    def test_recurrence_refuses_axised_driver(self):
        with pytest.raises(ValueError, match="finite axis"):
            mo.recurrence(
                start=0.0, formula="{prev} + {income}",
                variables={"income": self._axised()}, periods=4,
            )

    def test_sum_refuses(self):
        with pytest.raises(ValueError, match="finite axis"):
            mo.SUM(self._axised())

    def test_if_refuses(self):
        with pytest.raises(ValueError, match="finite axis"):
            mo.IF(self._axised() > 15, 1.0, 0.0)

    def test_plain_series_all_still_work(self):
        # The guard must not touch ordinary time series.
        m = mo.Model('m', default_grain='month', default_start='2025-01',
                     default_periods=4)
        with m:
            m.x = mo.Variable([1.0, 2.0, 3.0, 4.0], regrain=mo.up('sum'))
            m.s = mo.SUM(m.x)
            m.c = mo.cumsum(m.x)
        assert m.s._value == 10.0
        assert m.c._value == [1.0, 3.0, 6.0, 10.0]
        # Calendar quarters: Q1 = Jan+Feb+Mar, Q2 = Apr.
        assert m.x.at('quarter')._value == [6.0, 4.0]


class TestRank2EmissionGate:
    """A multi-axis Variable refuses one-row emission and
    grain projection with teaching errors instead of painting its
    cells as consecutive periods."""

    def _cube(self):
        months = mo.Variable(['Jan', 'Feb'])
        scenarios = mo.Variable(['Bear', 'Bull'])
        return mo.Variable(
            [1.0, 2.0, 3.0, 4.0], indexed_by=[months, scenarios],
            display_name='Cube',
        )

    def test_layout_refuses(self):
        from modeleon.compile.excel.layout import _reject_rank2_emission
        with pytest.raises(ValueError, match="2 axes"):
            _reject_rank2_emission(self._cube())

    def test_projection_refuses(self):
        from modeleon.core.projection import project_variable
        with pytest.raises(ValueError, match="axised"):
            project_variable(self._cube(), 'quarter', {})

    def test_rank1_untouched(self):
        from modeleon.compile.excel.layout import _reject_rank2_emission
        months = mo.Variable(['Jan', 'Feb'])
        _reject_rank2_emission(
            mo.Variable([1.0, 2.0], indexed_by=[months])
        )  # no raise


class TestKeysAlong:
    """``keys_along`` is the one predicate for "this line runs along keys,
    not periods": a keyed list, or a line on one finite non-time axis.
    The layout's ``line_keys`` is the same answer as text."""

    def test_keyed_list_gives_its_keys_as_given(self):
        from modeleon.core.shape import keys_along
        assert keys_along(mo.Variable({'Bear': 1, 'Base': 2})) == ['Bear', 'Base']
        byyear = mo.Variable([1.0, 2.0, 3.0], keys=[2026, 2027, 2028])
        assert keys_along(byyear) == [2026, 2027, 2028]

    def test_one_finite_axis_gives_its_labels(self):
        from modeleon.core.shape import keys_along
        scen = mo.Variable(['Bear', 'Bull'])
        assert keys_along(mo.Variable([1, 2], indexed_by=[scen])) == ['Bear', 'Bull']

    def test_none_for_a_series_a_scalar_a_time_axis_rank2_tracks_and_a_keyed_cell(self):
        from modeleon.core.shape import keys_along
        assert keys_along(mo.Variable([1, 2, 3])) is None
        assert keys_along(mo.Variable(7)) is None
        t = mo.Time('2026-01', periods=3)
        assert keys_along(mo.Variable([1, 2, 3], indexed_by=[t])) is None
        a = mo.Variable(['x', 'y'])
        b = mo.Variable(['p', 'q'])
        assert keys_along(mo.Variable([1, 2, 3, 4], indexed_by=[a, b])) is None
        mt = mo.Model('t', default_grain='month', default_start='2026-01', default_periods=3,
                      tracks=mo.Tracks('plan', 'actual'))
        mt.rev = mo.Variable(plan=[1, 2, 3], actual=[1, 2, 3])
        assert keys_along(mt.rev) is None
        assert keys_along(mo.Variable({'Bear': 1, 'Base': 2})['Base']) is None

    def test_layout_line_keys_is_keys_along_as_text(self):
        from modeleon.compile.excel.layout import line_keys
        from modeleon.core.shape import keys_along
        scen = mo.Variable(['Bear', 'Bull'])
        a = mo.Variable(['x', 'y'])
        b = mo.Variable(['p', 'q'])
        lines = [
            mo.Variable({'Bear': 1, 'Base': 2}),
            mo.Variable([1.0, 2.0, 3.0], keys=[2026, 2027, 2028]),
            mo.Variable([1, 2], indexed_by=[scen]),
            mo.Variable([1, 2, 3]),
            mo.Variable(7),
            mo.Variable([1, 2, 3], indexed_by=[mo.Time('2026-01', periods=3)]),
            mo.Variable([1, 2, 3, 4], indexed_by=[a, b]),
            mo.Variable({'Bear': 1, 'Base': 2})['Base'],
        ]
        for v in lines:
            raw = keys_along(v)
            assert line_keys(v) == (None if raw is None else [str(k) for k in raw])


def _keyed_model():
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3)
    m.growth = mo.Variable({'Bear': 0.9, 'Base': 1.0, 'Bull': 1.1})
    m.base = mo.Variable([100, 110, 120])
    return m


class TestKeysAgainstPeriods:
    """A line that runs along keys against a period series is refused:
    zipped cell by cell it pairs Bear with January — numbers that look
    right and are wrong. Everything the engine aligned on purpose keeps
    broadcasting as before."""

    def test_keyed_times_series_is_refused_naming_both_lines(self):
        m = _keyed_model()
        with pytest.raises(ValueError) as e:
            m.base * m.growth
        msg = str(e.value)
        assert "'growth' runs along its keys (Bear, Base, Bull)" in msg
        assert "'base' runs along periods (3 values)" in msg
        assert 'Bear with the first period, Base with the second' in msg
        assert 'Pick one key, growth["Bear"]' in msg
        assert "mo.Variable({'Bear': …, 'Base': …, 'Bull': …})" in msg

    def test_axised_times_series_is_refused_floating_and_named(self):
        scen = mo.Variable(['Bear', 'Base', 'Bull'])
        growth = mo.Variable([0.9, 1.0, 1.1], indexed_by=[scen])
        base = mo.Variable([100, 110, 120])
        with pytest.raises(ValueError) as e:
            base * growth
        msg = str(e.value)
        assert 'runs along an axis (Bear, Base, Bull)' in msg
        assert 'runs along periods (3 values)' in msg
        m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3)
        m.scen = mo.Variable(['Bear', 'Base', 'Bull'])
        m.g = mo.Variable([0.9, 1.0, 1.1], indexed_by=[m.scen], display_name='Growth by scenario')
        m.base = mo.Variable([100, 110, 120])
        with pytest.raises(ValueError) as e:
            m.base * m.g
        msg = str(e.value)
        assert "'Growth by scenario' runs along the axis 'scen' (Bear, Base, Bull)" in msg
        assert 'Pick one cell, g[0]' in msg
        assert 'indexed_by=[scen]' in msg

    @pytest.mark.parametrize('op', [
        operator.add, operator.sub, operator.mul, operator.truediv, operator.floordiv,
        operator.mod, operator.pow, operator.lt, operator.le, operator.gt, operator.ge,
        operator.eq, operator.ne,
    ])
    def test_every_operator_and_comparison_refuses(self, op):
        m = _keyed_model()
        with pytest.raises(ValueError, match='runs along'):
            op(m.growth, m.base)
        with pytest.raises(ValueError, match='runs along'):
            op(m.base, m.growth)

    def test_unequal_lengths_get_this_message_not_the_length_error(self):
        m = _keyed_model()
        with pytest.raises(ValueError) as e:
            m.growth * mo.Variable([1, 2, 3, 4])
        assert 'runs along periods (4 values)' in str(e.value)
        assert 'different lengths' not in str(e.value)

    def test_a_located_series_and_time_are_periods_too(self):
        m = _keyed_model()
        m.early = mo.Variable([1.0, 2.0, 3.0], start='2026-01', grain='month')
        with pytest.raises(ValueError, match="'early' runs along periods \\(3 values\\)"):
            m.growth * m.early
        with m:
            with pytest.raises(ValueError, match='the line mo.time.days runs along periods') as e:
                m.growth * mo.time.days
            # a time property cannot be given keys — the switch is offered
            assert 'switch with mo.CHOOSE(pick, growth)' in str(e.value)

    def test_operator_results_and_slices_still_run_along_keys(self):
        m = _keyed_model()
        with pytest.raises(ValueError, match=r'the line growth \* 2 runs along its keys'):
            (m.growth * 2) * m.base
        with pytest.raises(ValueError, match=r'runs along its keys \(Bear, Base\)'):
            m.growth[0:2] * m.base

    def test_year_keys_and_a_single_key_count(self):
        m = _keyed_model()
        m.byyear = mo.Variable([1.0, 2.0, 3.0], keys=[2026, 2027, 2028])
        with pytest.raises(ValueError) as e:
            m.byyear * m.base
        msg = str(e.value)
        assert 'runs along its keys (2026, 2027, 2028)' in msg
        assert 'Pick one cell, byyear[0]' in msg
        assert 'mo.Variable({2026: …, 2027: …, 2028: …})' in msg
        with pytest.raises(ValueError, match=r'runs along its keys \(Only\)'):
            mo.Variable({'Only': 5.0}) * mo.Variable([1, 2, 3])

    def test_scalars_and_one_value_lists_still_broadcast(self):
        m = _keyed_model()
        for r in (m.growth * 2, 2 * m.growth, m.growth * mo.Variable(2),
                  m.growth * mo.Variable([2])):
            assert r._value == pytest.approx([1.8, 2.0, 2.2])
            assert r._keys == ['Bear', 'Base', 'Bull']

    def test_a_raw_list_stays_positional(self):
        m = _keyed_model()
        r = m.growth * [1, 2, 3]
        assert r._value == pytest.approx([0.9, 2.0, 3.3])
        assert r._keys == ['Bear', 'Base', 'Bull']
        r2 = m.growth * m.base.value
        assert r2._value == pytest.approx([90.0, 110.0, 132.0])
        assert r2._keys == ['Bear', 'Base', 'Bull']

    def test_keyed_times_keyed_keeps_its_rules(self):
        m = _keyed_model()
        r = m.growth * mo.Variable({'Bull': 2, 'Base': 3, 'Bear': 4})
        assert r._value == pytest.approx([3.6, 3.0, 2.2])
        assert r._keys == ['Bear', 'Base', 'Bull']
        with pytest.raises(ValueError, match='different key sets'):
            m.growth * mo.Variable({'X': 1, 'Base': 2, 'Bear': 3})

    def test_a_keyed_cell_is_a_scalar_and_broadcasts(self):
        m = _keyed_model()
        r = m.growth['Base'] * m.base
        assert r._value == [100.0, 110.0, 120.0]
        assert r._keys is None

    def test_axes_and_the_time_axis_are_unchanged(self):
        t = mo.Time('2026-01', periods=3)
        rev = mo.Variable([100, 110, 120], indexed_by=[t])
        scen = mo.Variable(['Bear', 'Base', 'Bull'])
        growth = mo.Variable([0.9, 1.0, 1.1], indexed_by=[scen])
        r = mo.Variable([100, 110, 120]) * rev
        assert r._value == [10000, 12100, 14400]
        assert r._indexed_by == (t,)
        cross = rev * growth
        assert cross._indexed_by == (t, scen)
        assert len(cross._value) == 9
        same = growth * mo.Variable([1, 2, 3], indexed_by=[scen])
        assert same._value == pytest.approx([0.9, 2.0, 3.3])
        assert same._indexed_by == (scen,)

    def test_series_times_series_stays_positional(self):
        m = _keyed_model()
        assert (m.base * mo.Variable([1, 2, 3]))._value == [100, 220, 360]
        with m:
            days = (m.base * mo.time.days)._value
        assert days == [3100, 3080, 3720]

    def test_a_tracked_series_is_periods_too(self):
        mt = mo.Model('t', default_grain='month', default_start='2026-01', default_periods=3,
                      tracks=mo.Tracks('plan', 'actual'))
        mt.rev = mo.Variable(plan=[1, 2, 3], actual=[1, 2, 3])
        mt.g = mo.Variable({'Bear': 0.9, 'Base': 1.0, 'Bull': 1.1})
        for a, b in ((mt.rev, mt.g), (mt.g, mt.rev)):
            with pytest.raises(ValueError, match=r"'rev' runs along periods \(3 values\)"):
                a * b
        with pytest.raises(ValueError, match='runs along periods'):
            mt.rev > mt.g

    def test_a_series_on_the_time_axis_meets_a_keyed_list_and_is_refused(self):
        t = mo.Time('2026-01', periods=3)
        rev = mo.Variable([100, 110, 120], indexed_by=[t])
        g = mo.Variable({'Bear': 0.9, 'Base': 1.0, 'Bull': 1.1})
        with pytest.raises(ValueError, match='runs along periods'):
            g * rev

    def test_a_located_one_value_series_is_periods(self):
        m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=6)
        m.growth = mo.Variable({'Bear': 0.9, 'Base': 1.0, 'Bull': 1.1})
        m.fee = mo.Variable([100.0], start='2026-01', grain='month', extend=mo.zero())
        with pytest.raises(ValueError, match=r"'fee' runs along periods \(1 value\)"):
            m.fee * m.growth
        # a bare one-value list is still a single value
        assert (m.growth * mo.Variable([2.0]))._value == pytest.approx([1.8, 2.0, 2.2])

    def test_keyed_list_against_an_axis_of_other_labels_is_refused(self):
        regions = mo.Variable(['North', 'South', 'East'])
        by_region = mo.Variable([10, 20, 30], indexed_by=[regions])
        g = mo.Variable({'Bear': 0.9, 'Base': 1.0, 'Bull': 1.1})
        with pytest.raises(ValueError, match='labels would not match'):
            g * by_region
        scen = mo.Variable(['Bear', 'Base', 'Bull'])
        same = mo.Variable([1, 2, 3], indexed_by=[scen])
        assert (g * same)._value == pytest.approx([0.9, 2.0, 3.3])
        with pytest.raises(ValueError, match='labels would not match'):
            mo.Variable({'Bull': 1.1, 'Base': 1.0, 'Bear': 0.9}) * same

    def test_a_computed_series_side_gets_the_switch_as_the_fix(self):
        m = _keyed_model()
        m.doubled = m.base * 2
        with pytest.raises(ValueError) as e:
            m.growth * m.doubled
        assert 'switch with mo.CHOOSE(pick, growth)' in str(e.value)
        assert 'mo.Variable({' not in str(e.value)
