# SPDX-License-Identifier: Apache-2.0
"""Tests for MultiVariable - structural container."""

import pytest

from modeleon import MultiVariable, Variable


class TestMultiVariableFactory:
    def test_factory_style_registers_components(self):
        mv = MultiVariable(
            revenue=Variable(1_000_000),
            cogs=Variable(650_000),
        )
        assert "revenue" in mv._components
        assert "cogs" in mv._components
        assert mv._components["revenue"]._value == 1_000_000

    def test_component_access_via_attribute(self):
        mv = MultiVariable(revenue=Variable(1_000_000))
        assert mv.revenue._value == 1_000_000

    def test_display_name_kwarg(self):
        mv = MultiVariable(display_name="Income Statement", revenue=Variable(1))
        assert mv._display_name == "Income Statement"


class TestMultiVariableStrictKwargs:
    """``MultiVariable`` validates kwargs strictly — anything that isn't
    ``display_name``, ``excel_props``, or a Variable / MultiVariable
    component raises ``TypeError`` at construction. Mirrors
    :class:`Variable`'s strict-kwarg discipline."""

    def test_unknown_scalar_kwarg_raises(self):
        with pytest.raises(TypeError, match="unexpected keyword argument"):
            MultiVariable(row=5)

    def test_unknown_typo_kwarg_raises(self):
        with pytest.raises(TypeError, match="unexpected keyword argument"):
            MultiVariable(typo_name="pnl")

    def test_typo_with_typed_value_raises(self):
        with pytest.raises(TypeError, match="unexpected keyword argument"):
            MultiVariable(name=1)

    def test_excel_props_row_col_canonical_path(self):
        # ``row`` and ``col`` belong inside ``excel_props`` — that's the
        # canonical container layout reads from.
        mv = MultiVariable(excel_props={'row': 5, 'col': 2})
        assert mv.excel_props == {'row': 5, 'col': 2}


class TestMultiVariableImperative:
    def test_imperative_assignment_registers(self):
        mv = MultiVariable()
        mv.revenue = Variable(1_000_000)
        assert "revenue" in mv._components
        assert mv.revenue._value == 1_000_000


class TestMultiVariableContextManager:
    def test_with_block_is_alias(self):
        """``with mv as alias:`` is plain Python — alias is the same
        object. Attachment requires ``parent.child = ...``."""
        mv = MultiVariable(display_name="Costs")
        with mv as costs:
            assert costs is mv
            costs.cogs = Variable(650_000)
            costs.opex = Variable(200_000)
        assert "cogs" in mv._components
        assert "opex" in mv._components

    def test_attribute_assignment_preserves_creation_order(self):
        mv = MultiVariable()
        mv.a = Variable(1)
        mv.b = Variable(2)
        mv.c = Variable(3)
        assert mv._component_order == ["a", "b", "c"]


class TestSheetRole:
    def test_sheet_role_sets_is_sheet(self):
        s = MultiVariable("Income Statement", excel_props={'tab': True})
        assert s._is_sheet is True

    def test_sheet_factory(self):
        s = MultiVariable(
            "Assumptions",
            excel_props={'tab': True},
            tax_rate=Variable(0.25),
            cogs_pct=Variable(0.60),
        )
        assert "tax_rate" in s._components
        assert s.tax_rate._value == 0.25

    def test_sheet_with_block(self):
        s = MultiVariable("Income Statement", excel_props={'tab': True})
        s.revenue = Variable(1_000_000)
        s.cogs = s.revenue * Variable(0.6)
        assert "revenue" in s._components
        assert "cogs" in s._components
        assert s.cogs._value == 600_000


class TestNesting:
    def test_nested_multi_variables(self):
        tiers = MultiVariable(
            enterprise=MultiVariable(seats=Variable(100), price=Variable(200)),
            smb=MultiVariable(seats=Variable(50), price=Variable(80)),
        )
        assert tiers.enterprise.seats._value == 100
        assert tiers.smb.price._value == 80

    def test_cross_mv_reference_in_formula(self):
        assumptions = MultiVariable(tax_rate=Variable(0.25))
        pnl = MultiVariable()
        pnl.revenue = Variable(1_000_000)
        pnl.taxes = pnl.revenue * assumptions.tax_rate
        assert pnl.taxes._value == 250_000
