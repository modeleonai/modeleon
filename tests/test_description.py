# SPDX-License-Identifier: Apache-2.0
"""Tests for the ``description=`` kwarg + ``.description`` attribute on
:class:`Variable`, :class:`MultiVariable`, and :class:`Model`.

The field is free-form prose attached to a node by its author —
mirrors ``display_name`` as a structured surface but doesn't fall
back to anything. ``None`` when unset.
"""

from __future__ import annotations

import pytest

import modeleon as mo


class TestVariableDescription:
    def test_kwarg_sets_attribute(self):
        v = mo.Variable(100, description="Monthly revenue baseline.")
        assert v.description == "Monthly revenue baseline."

    def test_default_is_none(self):
        assert mo.Variable(100).description is None

    def test_chaining_setter(self):
        v = mo.Variable(50).set_description("after the fact")
        assert v.description == "after the fact"

    def test_setter_returns_self(self):
        v = mo.Variable(50)
        assert v.set_description("x") is v

    def test_set_description_none_clears(self):
        v = mo.Variable(1, description="initial").set_description(None)
        assert v.description is None

    def test_non_string_coerced(self):
        # Defensive: numeric / non-string descriptions get stringified.
        v = mo.Variable(1, description=42)
        assert v.description == "42"

    def test_independent_of_display_name(self):
        v = mo.Variable(1, display_name="Revenue", description="What I am")
        assert v.display_name == "Revenue"
        assert v.description == "What I am"


class TestMultiVariableDescription:
    def test_kwarg_via_positional(self):
        # MultiVariable's first positional is display_name; description
        # is a keyword-only arg in practice.
        mv = mo.MultiVariable("Assumptions", description="Inputs go here.")
        assert mv.description == "Inputs go here."

    def test_default_is_none(self):
        assert mo.MultiVariable("X").description is None

    def test_factory_kwargs_dont_swallow_description(self):
        # Factory style — child kwargs and description coexist.
        rate = mo.Variable(0.05)
        mv = mo.MultiVariable(
            "Assumptions",
            description="See deck p.12.",
            rate=rate,
        )
        assert mv.description == "See deck p.12."
        assert mv.rate is rate

    def test_chaining_setter_on_mv(self):
        mv = mo.MultiVariable("X").set_description("late binding")
        assert mv.description == "late binding"


class TestModelDescription:
    def test_kwarg(self):
        m = mo.Model("forecast", description="Q1 2026 baseline.")
        assert m.description == "Q1 2026 baseline."

    def test_default_is_none(self):
        assert mo.Model("forecast").description is None

    def test_description_on_root_visible_through_with(self):
        with mo.Model("forecast", description="root prose") as m:
            m.revenue = mo.Variable(100, description="opener")
        assert m.description == "root prose"
        assert m.revenue.description == "opener"


class TestUnknownKwargRejection:
    def test_mv_unknown_kwarg_still_rejects(self):
        # Adding description to the known-kwarg list shouldn't loosen
        # the typo guard. An unrelated kwarg still raises.
        with pytest.raises(TypeError, match="unexpected keyword"):
            from modeleon.core.multi_variable import MultiVariableBase
            MultiVariableBase(typo_kwarg=1)
