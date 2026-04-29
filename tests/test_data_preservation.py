# SPDX-License-Identifier: Apache-2.0
"""Tests for native data storage and the render_at projection.

A Variable carries native-form data in ``_data`` — keyed by the
authoring identity (integer position for plain lists, string labels
for keyed Variables). The ``render_at(period_idx)`` method projects
native data into the current view's period index.

Native storage is the foundation for axis-mutation safety: when a
Variable is bound to a future mutable provider (Calendar etc.) and
the provider's structure changes, the renderer reads through
``render_at`` against the still-authoritative ``_data`` rather than
re-deriving from a stale list.
"""

from __future__ import annotations

import modeleon as mo


class TestNativeDataPopulation:
    """Native data is captured at construction for list and dict
    inputs; scalars and formula-mode Variables skip it (their
    ``_value`` is already the source of truth)."""

    def test_scalar_has_no_native_data(self):
        v = mo.Variable(42)
        assert v._data is None

    def test_list_value_populates_positional_native_data(self):
        v = mo.Variable([100, 200, 300])
        assert v._data == {0: 100, 1: 200, 2: 300}

    def test_dict_value_populates_keyed_native_data(self):
        v = mo.Variable({"US": 100, "EU": 200, "APAC": 150})
        assert v._data == {"US": 100, "EU": 200, "APAC": 150}

    def test_explicit_keys_populate_keyed_native_data(self):
        v = mo.Variable([10, 20, 30], keys=["a", "b", "c"])
        assert v._data == {"a": 10, "b": 20, "c": 30}

    def test_formula_skips_native_data(self):
        a = mo.Variable(2)
        b = mo.Variable(3)
        c = a + b
        # Arithmetic intermediate — its _value is set by the operator
        # post-init. Native data isn't populated for derived results.
        assert c._data is None

    def test_recurrence_intermediate_skips_native_data(self):
        result = mo.recurrence(start=100, formula="{prev} * 1.05", periods=3)
        # Recurrence builds the list via _value assignment; native data
        # is reserved for user-authored input, not derived results.
        assert result._data is None


class TestRenderAt:
    """``render_at(period_idx)`` projects native data (or falls back to
    ``_value``) into the value at the given view-position."""

    def test_scalar_returns_scalar(self):
        v = mo.Variable(42)
        assert v.render_at(0) == 42

    def test_list_returns_value_at_index(self):
        v = mo.Variable([100, 200, 300])
        assert v.render_at(0) == 100
        assert v.render_at(1) == 200
        assert v.render_at(2) == 300

    def test_keyed_returns_value_at_keys_index(self):
        v = mo.Variable({"US": 10, "EU": 20, "APAC": 30})
        assert v.render_at(0) == 10
        assert v.render_at(1) == 20
        assert v.render_at(2) == 30

    def test_out_of_range_returns_none_for_lists(self):
        v = mo.Variable([1, 2, 3])
        assert v.render_at(99) is None

    def test_arithmetic_intermediate_renders_via_value_fallback(self):
        # Derived Variables don't have ``_data`` populated; render_at
        # falls back to ``_value`` so callers see consistent semantics.
        a = mo.Variable([1, 2, 3])
        b = mo.Variable([10, 20, 30])
        c = a + b
        assert c.render_at(0) == 11
        assert c.render_at(1) == 22
        assert c.render_at(2) == 33


class TestNativeDataIsAuthoritative:
    """Native data preserves the user's authored values regardless of
    later view operations on ``_value``."""

    def test_native_data_outlives_value_mutation(self):
        v = mo.Variable([100, 200, 300])
        assert v._data == {0: 100, 1: 200, 2: 300}
        # Simulate a view-side mutation (a future axis change might
        # rebuild ``_value`` shorter or longer). Native data stays.
        v._value = [100, 200]  # shrunk view
        assert v._data == {0: 100, 1: 200, 2: 300}
        # render_at still resolves to the authoritative native entry.
        assert v.render_at(2) == 300

    def test_native_data_indexing_independent_of_value_length(self):
        v = mo.Variable([1, 2, 3, 4, 5])
        v._value = []  # extreme: empty view
        # Native lookup still works.
        assert v.render_at(2) == 3
        assert v.render_at(4) == 5


class TestRoundTripThroughKeys:
    """For keyed Variables, render_at(i) must map through keys[i] —
    pinning the contract that integer view-position resolves to the
    keyed entry at that position."""

    def test_keyed_render_at_matches_dict_lookup(self):
        v = mo.Variable({"a": 1, "b": 2, "c": 3})
        for i, key in enumerate(["a", "b", "c"]):
            assert v.render_at(i) == v._data[key]

    def test_explicit_keys_round_trip(self):
        v = mo.Variable([10, 20, 30], keys=["x", "y", "z"])
        for i, key in enumerate(["x", "y", "z"]):
            assert v.render_at(i) == v._data[key]
