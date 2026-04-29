# SPDX-License-Identifier: Apache-2.0
"""Tests for the :class:`Component` typing layer.

Component is a thin marker subclass of :class:`Base` — it indicates
"this node lives in the model tree." Variable and MultiVariable both
inherit from Component. Metadata-only objects with identity but no
model-tree presence subclass Base directly. These tests pin that
contract.
"""

from __future__ import annotations

import modeleon as mo
from modeleon.core.base import Base
from modeleon.core.component import Component


class TestComponentInheritance:
    def test_component_is_a_base(self):
        assert issubclass(Component, Base)

    def test_variable_is_a_component(self):
        v = mo.Variable(1)
        assert isinstance(v, Component)

    def test_multivariable_is_a_component(self):
        m = mo.MultiVariable("Tab")
        assert isinstance(m, Component)


class TestComponentTypingTarget:
    """Functions that should accept any model-tree node take ``Component``
    as their parameter type. This pins that Variable and MultiVariable
    both satisfy that contract."""

    @staticmethod
    def takes_component(node: Component) -> str:
        return node.id

    def test_variable_passes_component_type(self):
        v = mo.Variable(1)
        # Should not raise; Variable is-a Component.
        assert isinstance(self.takes_component(v), str)

    def test_multivariable_passes_component_type(self):
        m = mo.MultiVariable("Tab")
        assert isinstance(self.takes_component(m), str)


class TestComponentExcelProps:
    """``excel_props``, ``set_style``, and the validation pattern are
    inherited from Component. Per-class ``_EXCEL_PROP_KEYS`` overrides
    determine what's accepted."""

    def test_variable_excel_props_default_empty(self):
        v = mo.Variable(1)
        assert v.excel_props == {}

    def test_multivariable_excel_props_default_empty(self):
        m = mo.MultiVariable()
        assert m.excel_props == {}

    def test_variable_accepts_known_excel_prop(self):
        v = mo.Variable(1, excel_props={"bold": True, "bg": "#fff"})
        assert v.excel_props == {"bold": True, "bg": "#fff"}

    def test_variable_rejects_unknown_excel_prop(self):
        import pytest
        with pytest.raises(TypeError, match="unknown key"):
            mo.Variable(1, excel_props={"bogus_key": True})

    def test_multivariable_accepts_tab_key(self):
        # ``tab`` is in MV's _EXCEL_PROP_KEYS, not in Variable's.
        m = mo.MultiVariable(excel_props={"tab": True})
        assert m.excel_props["tab"] is True

    def test_variable_rejects_tab_key(self):
        # ``tab`` is structural for MV; rejected on Variable.
        import pytest
        with pytest.raises(TypeError, match="unknown key"):
            mo.Variable(1, excel_props={"tab": True})

    def test_set_style_chains(self):
        v = mo.Variable(1)
        result = v.set_style(bold=True, bg="#fff")
        assert result is v
        assert v.excel_props["bold"] is True
        assert v.excel_props["bg"] == "#fff"

    def test_set_style_silently_ignores_unknown(self):
        v = mo.Variable(1)
        v.set_style(bogus=True, bold=True)
        assert "bogus" not in v.excel_props
        assert v.excel_props["bold"] is True


class TestComponentSetters:
    """``set_display_name`` and ``set_python_name`` are inherited from
    Component and chain ``self``."""

    def test_set_display_name_chains_on_variable(self):
        v = mo.Variable(1)
        result = v.set_display_name("Revenue")
        assert result is v
        assert v._display_name == "Revenue"

    def test_set_display_name_chains_on_multivariable(self):
        m = mo.MultiVariable()
        result = m.set_display_name("P&L")
        assert result is m
        assert m._display_name == "P&L"

    def test_set_python_name_chains_on_variable(self):
        v = mo.Variable(1)
        result = v.set_python_name("revenue")
        assert result is v
        assert v.python_name == "revenue"

    def test_set_python_name_chains_on_multivariable(self):
        m = mo.MultiVariable()
        result = m.set_python_name("pnl")
        assert result is m
        assert m.python_name == "pnl"

    def test_chaining_preserves_concrete_type(self):
        """``Self`` typing on chaining methods means the returned object
        retains the subclass type — chained calls keep access to
        subclass-specific attributes (e.g. ``Variable.value``)."""
        v = (
            mo.Variable(1_000_000)
            .set_display_name("Revenue")
            .set_style(bold=True)
            .set_python_name("revenue")
        )
        # If ``Self`` typing is wrong, the chain returns ``Component``
        # and ``.value`` (Variable-only) wouldn't be type-resolvable.
        # Runtime check that the final type is still Variable:
        assert isinstance(v, mo.Variable)
        assert v.value == 1_000_000


class TestBaseOnlyDoesNotPassComponentCheck:
    """Subclassing Base directly (without going through Component) yields
    a Base-but-not-Component object — the right fit for metadata-only
    types that aren't part of the model tree."""

    def test_bare_base_subclass_is_not_component(self):
        from modeleon.core.qpath import QPath

        class MetaNode(Base):
            @property
            def path(self) -> QPath:
                return QPath.floating(id(self), kind="m")

        obj = MetaNode(display_name="meta")
        assert isinstance(obj, Base)
        assert not isinstance(obj, Component)
