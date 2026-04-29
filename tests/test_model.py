# SPDX-License-Identifier: Apache-2.0
"""Tests for :class:`Model` — the canonical top-level container.

A :class:`Model` is a :class:`MultiVariable` with one extra contract:
its ``name`` positional crystallizes the root path at construction.
Children adopted under it inherit the rooted prefix automatically.
"""

from __future__ import annotations

import pytest

import modeleon as mo
from modeleon.core.qpath import QPath


class TestModelConstruction:
    def test_name_required(self):
        with pytest.raises(TypeError, match="missing 1 required positional argument"):
            mo.Model()  # type: ignore[call-arg]

    def test_name_must_be_string(self):
        with pytest.raises(TypeError, match="non-empty string"):
            mo.Model(42)  # type: ignore[arg-type]

    def test_empty_name_rejected(self):
        with pytest.raises(TypeError, match="non-empty string"):
            mo.Model("")

    def test_basic_construction(self):
        m = mo.Model("acme")
        assert m.python_name == "acme"
        assert m.id == "acme"
        assert not m.path.is_floating

    def test_display_name_defaults_to_humanized_name(self):
        m = mo.Model("income_statement")
        assert m.display_name == "Income Statement"

    def test_explicit_display_name_overrides_default(self):
        m = mo.Model("acme", display_name="Acme Corp")
        assert m.display_name == "Acme Corp"
        assert m.python_name == "acme"

    def test_excel_props_forwarded(self):
        m = mo.Model("acme", excel_props={"tab": True})
        assert m.excel_props == {"tab": True}

    def test_factory_components_forwarded(self):
        m = mo.Model(
            "quick",
            revenue=mo.Variable(1_000_000),
            cogs=mo.Variable(0.6),
        )
        assert "revenue" in m._components
        assert "cogs" in m._components
        assert m.revenue._value == 1_000_000


class TestModelInheritance:
    def test_model_is_multivariable(self):
        m = mo.Model("m")
        assert isinstance(m, mo.MultiVariable)

    def test_model_supports_adoption(self):
        m = mo.Model("m")
        m.x = mo.Variable(1)
        assert "x" in m._components
        assert m.x._owner is m

    def test_model_supports_with_block(self):
        with mo.Model("m") as model:
            model.x = mo.Variable(1)
        assert model.x._value == 1


class TestModelPathCrystallization:
    """Adopting under a Model gives every descendant a rooted path
    automatically — the model name is the only explicit naming gesture."""

    def test_top_level_path_rooted(self):
        m = mo.Model("acme")
        assert m.path == QPath(("acme",))

    def test_child_path_under_model(self):
        m = mo.Model("acme")
        m.pnl = mo.MultiVariable(excel_props={"tab": True})
        assert m.pnl.path == QPath(("acme", "pnl"))

    def test_grandchild_path_under_model(self):
        m = mo.Model("acme")
        m.pnl = mo.MultiVariable(excel_props={"tab": True})
        m.pnl.revenue = mo.Variable(1_000_000)
        assert m.pnl.revenue.path == QPath(("acme", "pnl", "revenue"))

    def test_two_models_have_independent_trees(self):
        a = mo.Model("a")
        a.x = mo.Variable(1)
        b = mo.Model("b")
        b.x = mo.Variable(2)
        assert a.x.id == "a.x"
        assert b.x.id == "b.x"
        assert a.x is not b.x

    def test_factory_children_are_path_rooted(self):
        m = mo.Model("acme", revenue=mo.Variable(1_000_000))
        assert m.revenue.path == QPath(("acme", "revenue"))


class TestModelDemotionOnAdoption:
    """A :class:`Model` is the root of a model tree by contract. Adopting
    one into another container ends its root-ness — its runtime class
    is rewritten to :class:`MultiVariable`. The user's reference still
    points to the same object; only the type marker changes."""

    def test_adoption_demotes_to_multivariable(self):
        a = mo.Model("a")
        assert isinstance(a, mo.Model)
        t = mo.Model("t")
        t.a = a
        # Same object — but its class is now MultiVariable, not Model.
        assert t.a is a
        assert not isinstance(a, mo.Model)
        assert isinstance(a, mo.MultiVariable)

    def test_adoption_under_multivariable_demotes(self):
        sub = mo.Model("sub")
        container = mo.MultiVariable()
        container.sub = sub
        assert not isinstance(sub, mo.Model)
        assert isinstance(sub, mo.MultiVariable)

    def test_demoted_path_inherits_parent_root(self):
        a = mo.Model("a")
        a.x = mo.Variable(1)   # rooted at 'a' before demotion
        assert a.x.path == QPath(("a", "x"))

        t = mo.Model("t")
        t.a = a
        # After demotion + adoption, ``a`` is a sub-tree of ``t``;
        # its descendants reroot under the new prefix.
        assert a.path == QPath(("t", "a"))
        assert a.x.path == QPath(("t", "a", "x"))

    def test_factory_pattern_demotes_model_components(self):
        sub = mo.Model("sub")
        outer = mo.MultiVariable(sub=sub)
        assert outer.sub is sub
        assert not isinstance(sub, mo.Model)
        assert isinstance(sub, mo.MultiVariable)

    def test_demoted_model_can_be_adopted_again_as_multivariable(self):
        # Once demoted, the node behaves like any other MV — it can
        # be re-adopted, replaced, etc. without the root-only check
        # firing again.
        a = mo.Model("a")
        t1 = mo.Model("t1")
        t1.a = a   # demotes
        t2 = mo.Model("t2")
        # Cross-parent adoption — clones rather than moves.
        t2.a = a
        assert t2.a.path == QPath(("t2", "a"))

    def test_demotion_does_not_mutate_excel_props(self):
        # Demotion is a class swap only; explicit Excel markers like
        # ``tab`` aren't auto-set. Sheet-vs-section is decided
        # positionally by the renderer, not baked on the object.
        acme = mo.Model("acme")
        t = mo.Model("t")
        t.h = acme
        assert not acme._is_sheet
        assert acme.excel_props == {}
