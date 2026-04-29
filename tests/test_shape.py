# SPDX-License-Identifier: Apache-2.0
"""Tests for the :class:`Shape` attribute on Variable.

Every Variable carries exactly one ``shape``. The tests here pin:

- The default scalar Shape for a plain ``Variable(scalar)``.
- Auto-inference of an anonymous :class:`PositionalAxis` for list
  values.
- Explicit ``axis=`` and ``time_behavior=`` kwargs flow through.
- ``shape.length`` matches the cell count regardless of how the axis
  was constructed.
- Live-binding semantics — ``shape.axis`` reads through the bound
  owner so mutating the owner is observed by every Variable bound
  to it.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import List

import modeleon as mo
from modeleon.core.axis import Axis, PositionalAxis
from modeleon.core.shape import (
    AxisProvider,
    MutableAxisProvider,
    Shape,
    _AnonymousPositionalProvider,
)


class TestDefaultShape:
    def test_scalar_variable_has_scalar_shape(self):
        v = mo.Variable(42)
        assert v.shape.is_scalar()
        assert v.shape.length == 1
        assert v.shape.axis is None

    def test_scalar_shape_has_no_behavior_by_default(self):
        v = mo.Variable(42)
        assert v.shape.behavior is None

    def test_scalar_shape_has_no_keys_by_default(self):
        v = mo.Variable(42)
        assert v.shape.keys is None


class TestListAutoInferShape:
    def test_list_value_gets_positional_axis(self):
        v = mo.Variable([100, 200, 300])
        assert isinstance(v.shape.axis, PositionalAxis)
        assert v.shape.length == 3

    def test_positional_axis_uses_value_length(self):
        v = mo.Variable([1, 2, 3, 4, 5])
        assert v.shape.length == 5
        assert v.shape.axis.length == 5

    def test_positional_axis_has_position_name(self):
        v = mo.Variable([1, 2, 3])
        assert v.shape.axis.name == "position"

    def test_list_shape_is_not_scalar(self):
        v = mo.Variable([1, 2, 3])
        assert not v.shape.is_scalar()


class TestExplicitAxisKwarg:
    def test_explicit_axis_provider_is_stored(self):
        provider = _AnonymousPositionalProvider(PositionalAxis(length=12, name="month"))
        v = mo.Variable([10] * 12, axis=provider)
        assert v.shape.axis_owner is provider
        assert v.shape.axis is provider.axis
        assert v.shape.axis.name == "month"
        assert v.shape.length == 12

    def test_time_behavior_kwarg_is_stored(self):
        v = mo.Variable([10] * 3, time_behavior="flow")
        assert v.shape.behavior == "flow"

    def test_time_behavior_independent_of_axis(self):
        v = mo.Variable(1000, time_behavior="point")
        assert v.shape.is_scalar()
        assert v.shape.behavior == "point"


class TestKeyedShape:
    def test_dict_construction_populates_shape_keys(self):
        v = mo.Variable({"US": 100, "EU": 200, "APAC": 150})
        assert v.shape.keys == ("US", "EU", "APAC")

    def test_explicit_keys_kwarg_populates_shape_keys(self):
        v = mo.Variable([10, 20, 30], keys=["a", "b", "c"])
        assert v.shape.keys == ("a", "b", "c")


class TestLiveAxisBinding:
    """``Shape.axis_owner`` is a *reference* — mutating the provider's
    axis is observed by every Variable bound to it on next read."""

    def test_dereference_reads_current_axis(self):
        @dataclass
        class MutableProvider:
            axis: Axis

            def _invalidate_dependents(self) -> None:
                return None

        provider = MutableProvider(axis=Axis(length=5, name="t"))
        v = mo.Variable([1, 2, 3, 4, 5], axis=provider)
        assert v.shape.length == 5

        # Mutate the provider's axis. Variable observes the change live.
        provider.axis = Axis(length=10, name="t")
        assert v.shape.length == 10

    def test_two_variables_share_axis_by_identity(self):
        provider = _AnonymousPositionalProvider(PositionalAxis(length=4))
        a = mo.Variable([1, 2, 3, 4], axis=provider)
        b = mo.Variable([10, 20, 30, 40], axis=provider)
        assert a.shape.axis_owner is b.shape.axis_owner

    def test_axis_provider_protocol_runtime_check(self):
        @dataclass
        class CustomProvider:
            axis: Axis = field(default_factory=lambda: Axis(length=3))

            def _invalidate_dependents(self) -> None:
                return None

        provider = CustomProvider()
        assert isinstance(provider, AxisProvider)


class TestMutableAxisProvider:
    """``MutableAxisProvider`` tracks bound Variables in a weak set and
    broadcasts an invalidation callback on every ``adjust(...)``.

    Concrete time-axis providers (calendars, planning horizons, etc.)
    subclass this and let users mutate the axis at runtime — bound
    Variables observe the new state on next read."""

    def test_adjust_replaces_axis_atomically(self):
        provider = MutableAxisProvider(Axis(length=3, name="t"))
        original = provider.axis
        provider.adjust(length=10)
        assert provider.axis is not original
        assert provider.axis.length == 10
        assert provider.axis.name == "t"  # unchanged

    def test_bound_variable_observes_axis_length_change(self):
        provider = MutableAxisProvider(PositionalAxis(length=3))
        v = mo.Variable([1, 2, 3], axis=provider)
        assert v.shape.length == 3
        provider.adjust(length=12)
        assert v.shape.length == 12

    def test_dependent_invalidation_fires_per_variable(self):
        provider = MutableAxisProvider(Axis(length=3))
        invalidations: List[str] = []

        class TrackingVariable(mo.Variable):
            def _on_axis_invalidated(self) -> None:
                invalidations.append(self.id)

        a = TrackingVariable([1, 2, 3], axis=provider, display_name="A")
        b = TrackingVariable([10, 20, 30], axis=provider, display_name="B")
        a._python_name = "a"
        b._python_name = "b"

        provider.adjust(length=4)
        assert sorted(invalidations) == sorted([a.id, b.id])

    def test_two_adjustments_fire_invalidation_twice(self):
        provider = MutableAxisProvider(Axis(length=3))
        count = 0

        class CountingVariable(mo.Variable):
            def _on_axis_invalidated(self) -> None:
                nonlocal count
                count += 1

        v = CountingVariable([1, 2, 3], axis=provider)
        provider.adjust(length=4)
        provider.adjust(length=5)
        assert count == 2

    def test_dropped_variable_does_not_keep_provider_alive(self):
        import gc
        provider = MutableAxisProvider(Axis(length=3))
        v = mo.Variable([1, 2, 3], axis=provider)
        # ``v`` is the only strong reference — dropping it removes it
        # from the provider's weak set after garbage collection.
        del v
        gc.collect()
        assert len(provider._dependents) == 0

    def test_provider_satisfies_axis_provider_protocol(self):
        provider = MutableAxisProvider(Axis(length=3))
        assert isinstance(provider, AxisProvider)


class TestExistingBehaviorUnchanged:
    """``shape`` is advisory in this introduction — layout, arithmetic,
    and renderers still operate on ``_value`` directly. These tests pin
    that no existing path changed shape on construction."""

    def test_arithmetic_intermediate_has_a_shape(self):
        a = mo.Variable([1, 2, 3])
        b = mo.Variable([10, 20, 30])
        c = a + b
        # Result is a list-shaped Variable.
        assert c.shape.length == 3

    def test_scalar_plus_scalar_stays_scalar(self):
        a = mo.Variable(2)
        b = mo.Variable(3)
        c = a + b
        assert c.shape.is_scalar()
