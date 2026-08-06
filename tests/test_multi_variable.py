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


class TestAdd:
    """``mv.add(name, component)`` — the dynamic-name form of attribute
    assignment (runtime-computed child names)."""

    def test_add_variable_by_runtime_name(self):
        mv = MultiVariable("P&L")
        name = "rev" + "_2026"
        got = mv.add(name, Variable(100))
        assert "rev_2026" in mv._components
        assert mv.rev_2026._value == 100
        assert got is mv.rev_2026

    def test_add_multivariable(self):
        parent = MultiVariable("root")
        child = parent.add("board", MultiVariable("Board"))
        assert "board" in parent._components
        assert parent.board is child
        # Namespace flows through — the child's qpath descends from parent.
        parent.board.note = Variable(1)
        assert parent.board.note._value == 1

    def test_add_is_equivalent_to_setattr(self):
        a = MultiVariable("a")
        a.x = Variable(5)
        b = MultiVariable("b")
        b.add("x", Variable(5))
        assert a.x._value == b.x._value == 5
        assert a._component_names == b._component_names

    def test_add_clones_on_reparent_and_returns_clone(self):
        # A Variable already owned elsewhere is cloned on adoption; add
        # must return the adopted (clone) instance, not the original.
        src = MultiVariable("src")
        src.v = Variable(7)
        dst = MultiVariable("dst")
        got = dst.add("v", src.v)
        assert dst.v._value == 7
        assert got is dst.v

    def test_add_rejects_private_name(self):
        mv = MultiVariable()
        with pytest.raises(ValueError):
            mv.add("_secret", Variable(1))

    def test_add_rejects_empty_name(self):
        mv = MultiVariable()
        with pytest.raises(ValueError):
            mv.add("", Variable(1))

    def test_add_rejects_non_component(self):
        mv = MultiVariable()
        with pytest.raises(TypeError):
            mv.add("f", lambda: 1)


class TestRemove:
    """``mv.remove(name)`` — detach a child; inverse of add/assignment."""

    def test_remove_detaches_child(self):
        mv = MultiVariable("root")
        mv.a = Variable(1)
        mv.b = Variable(2)
        got = mv.remove("a")
        assert "a" not in mv._components
        assert "b" in mv._components
        assert got._value == 1  # returned the detached instance
        assert not hasattr(mv, "a")

    def test_remove_missing_raises_keyerror(self):
        mv = MultiVariable()
        with pytest.raises(KeyError):
            mv.remove("nope")

    def test_remove_then_readd_same_name(self):
        mv = MultiVariable()
        mv.x = Variable(1)
        mv.remove("x")
        mv.add("x", Variable(9))  # no orphan warning — the slot was freed
        assert mv.x._value == 9

    def test_remove_nested_mv(self):
        parent = MultiVariable("p")
        parent.child = MultiVariable("c")
        parent.child.leaf = Variable(5)
        detached = parent.remove("child")
        assert "child" not in parent._components
        assert detached.leaf._value == 5  # subtree intact, just detached


class TestMultiVariableClassComputeParams:
    """Constructor kwargs must reach compute() — including required
    (no-default) parameters, the documented template style."""

    def test_required_params_flow_from_kwargs(self):
        from modeleon.core.multi_variable import MultiVariableClass

        class Unit(MultiVariableClass):
            def compute(self, seats, price):
                self.revenue = Variable(seats) * Variable(price)

        u = Unit(seats=100, price=10)
        assert u.revenue._value == 1000

    def test_variable_param_preserves_formula_tracking(self):
        from modeleon.core.multi_variable import MultiVariableClass

        class Unit(MultiVariableClass):
            def compute(self, rate):
                self.doubled = rate * Variable(2)

        rate = Variable(0.05)
        u = Unit(rate=rate)
        # the Variable itself (not the extracted scalar) reached compute
        assert u.doubled._value == pytest.approx(0.1)
        assert u.doubled._expr is not None

    def test_default_still_applies_when_kwarg_omitted(self):
        from modeleon.core.multi_variable import MultiVariableClass

        class Unit(MultiVariableClass):
            def compute(self, base, factor=3):
                self.out = Variable(base) * Variable(factor)

        u = Unit(base=5)
        assert u.out._value == 15

    def test_missing_required_param_raises_named_typeerror(self):
        from modeleon.core.multi_variable import MultiVariableClass

        class Unit(MultiVariableClass):
            def compute(self, seats):
                self.x = Variable(seats)

        with pytest.raises(TypeError, match="seats"):
            Unit()


class TestMultiVariableClassShell:
    """__shell__ constructs a real instance without invoking compute() —
    the container starts empty and is populated explicitly by the caller
    (deserialization / code-generation flows)."""

    def _unit_cls(self):
        from modeleon.core.multi_variable import MultiVariableClass

        class Unit(MultiVariableClass):
            def compute(self, seats, price):
                self.revenue = Variable(seats) * Variable(price)

            def describe(self):
                return f"unit:{self.seats}"

        return Unit

    def test_shell_skips_compute_but_keeps_identity(self):
        Unit = self._unit_cls()
        u = Unit.__shell__(seats=100, price=10)
        assert isinstance(u, Unit)          # real class, real MRO
        assert "revenue" not in u._components  # compute never ran
        assert u.describe() == "unit:100"   # methods work
        assert u._compute_suppressed is True

    def test_shell_stores_params_like_normal_construction(self):
        Unit = self._unit_cls()
        rate = Variable(0.05)
        u = Unit.__shell__(seats=100, price=10, churn=rate)
        assert u.seats == 100
        assert u._input_variables["churn"] is rate

    def test_shell_container_populates_explicitly(self):
        Unit = self._unit_cls()
        u = Unit.__shell__(seats=100, price=10)
        u.revenue = Variable(100) * Variable(10)
        assert u.revenue._value == 1000
        assert "revenue" in u._components

    def test_normal_construction_still_computes(self):
        Unit = self._unit_cls()
        u = Unit(seats=100, price=10)
        assert u.revenue._value == 1000
        assert u._compute_suppressed is False

    def test_argument_position_constructors_still_compute(self):
        # Other(...) evaluates BEFORE __shell__ enters its scope — it
        # must compute normally even when passed into a shell call.
        from modeleon.core.multi_variable import MultiVariableClass
        Unit = self._unit_cls()

        class Other(MultiVariableClass):
            def compute(self, base):
                self.out = Variable(base) * Variable(2)

        u = Unit.__shell__(seats=1, price=1, sub=Other(base=5))
        assert u.sub.out._value == 10  # inner compute ran

    def test_shell_then_normal_isolated(self):
        # Suppression must not leak past __shell__'s scope.
        Unit = self._unit_cls()
        Unit.__shell__(seats=1, price=1)
        u = Unit(seats=3, price=4)
        assert u.revenue._value == 12
