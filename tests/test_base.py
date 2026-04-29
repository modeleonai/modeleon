# SPDX-License-Identifier: Apache-2.0
"""Tests for the shared :class:`Base` identity surface.

Variable and MultiVariable both inherit identity (``_qualified_id``,
``_python_name``, ``_display_name``, ``python_name``, ``id``) from
:class:`modeleon.core.base.Base` via
:class:`modeleon.core.component.Component`. These tests pin the
shared surface and the contract for direct ``Base`` subclasses.
"""

from __future__ import annotations

import pytest

import modeleon as mo
from modeleon.core.base import Base
from modeleon.core.component import Component
from modeleon.core.qpath import QPath


class TestBaseInheritance:
    """Variable and MultiVariable both reach Base via Component."""

    def test_variable_is_a_component(self):
        v = mo.Variable(1)
        assert isinstance(v, Component)
        assert isinstance(v, Base)

    def test_multivariable_is_a_component(self):
        m = mo.MultiVariable("Tab")
        assert isinstance(m, Component)
        assert isinstance(m, Base)


class TestBaseFields:
    """The three fields Base owns are initialized identically on
    Variable and MultiVariable construction."""

    def test_variable_has_qualified_id_field(self):
        v = mo.Variable(1)
        # Field exists; un-crystallized Variable starts with None.
        assert hasattr(v, "_qualified_id")
        assert v._qualified_id is None

    def test_multivariable_has_qualified_id_field(self):
        m = mo.MultiVariable("Tab")
        assert hasattr(m, "_qualified_id")
        assert m._qualified_id is None

    def test_python_name_starts_none(self):
        v = mo.Variable(1)
        m = mo.MultiVariable("Tab")
        assert v._python_name is None
        assert m._python_name is None

    def test_display_name_constructor_sets_underlying_field(self):
        v = mo.Variable(1, display_name="Revenue")
        m = mo.MultiVariable(display_name="Custom Label")
        assert v._display_name == "Revenue"
        assert m._display_name == "Custom Label"

    def test_display_name_none_keeps_underlying_none(self):
        v = mo.Variable(1)
        m = mo.MultiVariable()
        assert v._display_name is None
        assert m._display_name is None


class TestBasePythonNameReadOnly:
    """``python_name`` is a read-only property inherited from Base.
    It's set internally by adoption (``parent.x = node`` writes
    ``node._python_name = "x"``) or by ``Model('name')`` for the root.
    Public assignment is rejected."""

    def test_property_is_read_only(self):
        v = mo.Variable(1)
        with pytest.raises(AttributeError):
            v.python_name = "revenue"

    def test_property_is_read_only_on_mv(self):
        m = mo.MultiVariable("Tab")
        with pytest.raises(AttributeError):
            m.python_name = "pnl"

    def test_adoption_sets_python_name(self):
        m = mo.Model("m")
        m.x = mo.Variable(1)
        assert m.x.python_name == "x"

    def test_model_constructor_sets_python_name(self):
        acme = mo.Model("acme")
        assert acme.python_name == "acme"


class TestBaseIdProperty:
    """``id`` is ``str(self.path)`` and is inherited from Base."""

    def test_id_is_string(self):
        v = mo.Variable(1)
        assert isinstance(v.id, str)

    def test_id_matches_path(self):
        v = mo.Variable(1)
        m = mo.MultiVariable("Tab")
        assert v.id == str(v.path)
        assert m.id == str(m.path)

    def test_adopted_id_is_qualified(self):
        # ``Model`` is the canonical top-level entry — its ``name``
        # crystallizes the root path so adopted children inherit a
        # rooted prefix.
        pnl = mo.Model("pnl")
        pnl.revenue = mo.Variable(100)
        assert pnl.id == "pnl"
        assert pnl.revenue.id == "pnl.revenue"

    def test_top_level_mv_without_python_name_stays_floating(self):
        # ``with mo.MultiVariable() as pnl:`` aliases the local binding
        # but does not set ``_python_name``. The MV stays floating;
        # adopted Variables inherit a floating prefix.
        with mo.MultiVariable() as pnl:
            pnl.revenue = mo.Variable(100)
        assert pnl.path.is_floating
        assert pnl.revenue.path.is_floating
        assert pnl.revenue.python_name == "revenue"  # set by adoption
        assert pnl.python_name is None  # never set


class TestBaseSubclassDirectly:
    """Subclassing Base directly (without Component) should yield a node
    with identity + label but no model-tree composition. This is the
    contract for metadata-only types that aren't part of the model tree."""

    def test_bare_base_subclass_has_identity_fields(self):
        class MetaNode(Base):
            @property
            def path(self) -> QPath:
                return QPath.floating(id(self), kind="m")

        m = MetaNode(display_name="meta")
        assert m._qualified_id is None
        assert m._python_name is None
        assert m._display_name == "meta"
        assert m.python_name is None
        assert m.id == str(m.path)

    def test_path_not_implemented_on_bare_base(self):
        b = Base()
        with pytest.raises(NotImplementedError):
            _ = b.path
