# SPDX-License-Identifier: Apache-2.0
"""Qualified-path identity.

Every Variable / MultiVariable carries a ``.path: QPath`` from the
model root. Paths crystallize at adoption time (via attribute
assignment, the only attachment path) and stay stable across sessions.
Unattached objects get a synthetic ``__floating__.{kind}{id}`` path.
"""

import modeleon as mo
from modeleon.core.qpath import QPath


class TestFloatingIdentity:
    """Unattached Variables / MVs have synthetic floating paths."""

    def test_unattached_variable_is_floating(self):
        v = mo.Variable(1)
        assert v.path.is_floating

    def test_unattached_mv_is_floating(self):
        mv = mo.MultiVariable("X")
        assert mv.path.is_floating

    def test_two_unattached_variables_get_distinct_paths(self):
        a = mo.Variable(1)
        b = mo.Variable(2)
        assert a.path != b.path

    def test_python_name_set_explicitly_crystallizes(self):
        v = mo.Variable(1)
        v._python_name = 't'
        # An unattached Variable with a python_name resolves to ROOT.t
        assert v.id == 't'


class TestAttachmentSetsPath:
    """``parent.child = ...`` sets the child's python_name and path."""

    def test_attached_variable_has_path_under_parent(self):
        m = mo.MultiVariable("M")
        m._python_name = 'm'
        m.x = mo.Variable(1)
        assert m.x.id == 'm.x'
        assert m.x.python_name == 'x'

    def test_attached_sub_mv_has_path_under_parent(self):
        m = mo.MultiVariable("M")
        m._python_name = 'm'
        m.s = mo.MultiVariable("S", excel_props={'tab': True})
        assert m.s.id == 'm.s'
        assert m.s.python_name == 's'

    def test_deeply_nested_paths(self):
        m = mo.MultiVariable("M")
        m._python_name = 'm'
        m.s = mo.MultiVariable("S", excel_props={'tab': True})
        m.s.x = mo.Variable(10)
        assert m.s.x.id == 'm.s.x'

    def test_explicit_python_name_overrides_default(self):
        m = mo.MultiVariable("M")
        m._python_name = 'root'
        m.x = mo.Variable(1)
        m.x._python_name = 'override'
        # python_name set explicitly is honored.
        assert m.x.python_name == 'override'


class TestCrossSheetIdentity:
    """Variables with the same local name in different sheets have distinct paths."""

    def test_same_name_different_sheets_distinct_paths(self):
        m = mo.MultiVariable("M")
        m._python_name = 'm'
        m.a = mo.MultiVariable("A", excel_props={'tab': True})
        m.a.cogs = mo.Variable(100)
        m.b = mo.MultiVariable("B", excel_props={'tab': True})
        m.b.cogs = mo.Variable(200)
        assert m.a.cogs.id == 'm.a.cogs'
        assert m.b.cogs.id == 'm.b.cogs'
        assert m.a.cogs is not m.b.cogs


class TestQualifiedPath:
    """Rooted MV trees get qualified paths end-to-end."""

    def test_model_path_from_python_name(self):
        m = mo.MultiVariable("Acme")
        m._python_name = 'm'
        assert m.path == QPath(("m",))

    def test_sheet_path_under_model(self):
        m = mo.MultiVariable("Acme")
        m._python_name = 'm'
        m.pnl = mo.MultiVariable("P&L", excel_props={'tab': True})
        assert m.pnl.path == QPath(("m", "pnl"))

    def test_variable_path_under_sheet(self):
        m = mo.MultiVariable("Acme")
        m._python_name = 'm'
        m.pnl = mo.MultiVariable("P&L", excel_props={'tab': True})
        m.pnl.revenue = mo.Variable(1_000_000)
        assert m.pnl.revenue.path == QPath(("m", "pnl", "revenue"))

    def test_factory_style_components_crystallize(self):
        m = mo.MultiVariable("Forecast")
        m._python_name = 'm'
        m.assumptions = mo.MultiVariable(
            "Assumptions", excel_props={'tab': True},
            tax_rate=mo.Variable(0.25),
            cogs_pct=mo.Variable(0.6),
        )
        assert m.assumptions.path == QPath(("m", "assumptions"))
        assert m.assumptions.tax_rate.path == QPath(("m", "assumptions", "tax_rate"))
        assert m.assumptions.cogs_pct.path == QPath(("m", "assumptions", "cogs_pct"))


class TestCrossModelIsolation:
    """Two models built in the same process don't share identity state."""

    def test_two_models_have_independent_trees(self):
        a = mo.MultiVariable("A")
        a._python_name = 'a'
        a.x = mo.Variable(1)
        b = mo.MultiVariable("B")
        b._python_name = 'b'
        b.x = mo.Variable(2)

        assert a.x.id == 'a.x'
        assert b.x.id == 'b.x'
        assert a.x is not b.x


class TestDualKeyedAddresses:
    """Layout's address map is keyed by qualified path."""

    def test_addresses_keyed_by_path(self):
        from modeleon.compile.excel.layout import LayoutEngine

        m = mo.MultiVariable("M")
        m._python_name = 'm'
        m.s = mo.MultiVariable("S", excel_props={'tab': True})
        m.s.x = mo.Variable(1, display_name='X')
        engine = LayoutEngine([m.s])
        addresses = engine.compute_addresses()
        # Address resolves under the qualified path:
        assert m.s.x.id in addresses


class TestDependencyProjection:
    """``.dependencies`` projects ``_dependency_refs`` to current paths."""

    def test_dependencies_use_qualified_paths(self):
        m = mo.MultiVariable("M")
        m._python_name = 'm'
        m.a = mo.Variable(10)
        m.b = mo.Variable(20)
        m.c = m.a + m.b
        assert m.c.dependencies == {'m.a', 'm.b'}
