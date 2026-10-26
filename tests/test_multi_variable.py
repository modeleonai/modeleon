# SPDX-License-Identifier: Apache-2.0
"""Tests for MultiVariable - structural container."""

import warnings

import pytest

from modeleon import Model, MultiVariable, Variable
from modeleon.core.multi_variable import MultiVariableClass


def _workbook_cells(path):
    """Every sheet's cell values, row by row — enough to compare two
    written workbooks for a stray or missing row."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    return {ws.title: [[c.value for c in row] for row in ws.iter_rows()] for ws in wb}


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


class TestReservedComponentNames:
    """Names the node itself answers — ``code``, ``description``,
    ``excel_props``, ``python_name``, ``path``, ``id``, … — are
    read-only properties, so they cannot also name a component. The
    refusal comes before anything is registered: the container is left
    exactly as it was, and nothing extra reaches the workbook."""

    RESERVED = (
        "code", "description", "excel_props", "python_name",
        "path", "id", "component_names",
    )

    @pytest.mark.parametrize("name", RESERVED)
    def test_assignment_refused_without_side_effects(self, name):
        mv = MultiVariable()
        mv.a = Variable(1)
        before = mv.to_dict()
        with pytest.raises(AttributeError, match=name):
            setattr(mv, name, Variable(5))
        assert mv.component_names == ["a"]
        assert mv.to_dict() == before

    def test_add_with_reserved_name_refused(self):
        mv = MultiVariable()
        mv.a = Variable(1)
        with pytest.raises(AttributeError, match="code"):
            mv.add("code", Variable(5))
        assert mv.component_names == ["a"]

    def test_factory_with_reserved_name_refused_before_adopting(self):
        a = Variable(1)
        with pytest.raises(AttributeError, match="code"):
            MultiVariable(a=a, code=Variable(5))
        # nothing was adopted onto the refused container
        assert a._owner is None
        assert a.python_name is None

    def test_refused_model_is_left_a_model(self):
        inner = Model("inner")
        mv = MultiVariable()
        with pytest.raises(AttributeError):
            mv.code = inner
        assert type(inner) is Model
        assert inner._parent is None

    def test_workbook_has_no_stray_row(self, tmp_path):
        m = Model("m")
        m.s = MultiVariable(excel_props={"tab": True})
        m.s.a = Variable(1)
        m.to_excel(tmp_path / "before.xlsx")
        with pytest.raises(AttributeError):
            m.s.code = Variable(5)
        m.to_excel(tmp_path / "after.xlsx")
        assert _workbook_cells(tmp_path / "after.xlsx") == _workbook_cells(
            tmp_path / "before.xlsx"
        )

    def test_subclass_read_only_property_is_reserved_too(self):
        class Priced(MultiVariable):
            @property
            def total(self):
                return 42

        p = Priced()
        with pytest.raises(AttributeError, match="total"):
            p.total = Variable(1)
        assert p.component_names == []
        assert p.total == 42


class TestDelComponent:
    """``del mv.x`` on a component is the same detach as
    ``mv.remove('x')``; on any other attribute it is plain Python."""

    def test_del_detaches_like_remove(self):
        mv = MultiVariable("root")
        mv.a = Variable(1)
        mv.b = Variable(2)
        b = mv.b
        del mv.b
        assert mv.component_names == ["a"]
        assert not hasattr(mv, "b")
        assert b._owner is None
        assert b.python_name is None

    def test_del_nested_mv(self):
        parent = MultiVariable("p")
        parent.child = MultiVariable("c")
        parent.child.leaf = Variable(5)
        child = parent.child
        del parent.child
        assert parent.component_names == []
        assert child._parent is None
        assert child.leaf._value == 5  # subtree intact, just detached

    def test_to_excel_repr_and_to_dict_work_after_del(self, tmp_path):
        def build(with_b):
            m = Model("m")
            m.s = MultiVariable(excel_props={"tab": True})
            m.s.a = Variable(1)
            if with_b:
                m.s.b = Variable(2)
            return m

        m = build(with_b=True)
        del m.s.b
        m.to_excel(tmp_path / "deleted.xlsx")
        build(with_b=False).to_excel(tmp_path / "never.xlsx")
        assert _workbook_cells(tmp_path / "deleted.xlsx") == _workbook_cells(
            tmp_path / "never.xlsx"
        )
        assert m._repr_html_()
        assert m.s.to_dict()["components"] == ["a"]

    def test_readd_after_del_is_silent(self):
        mv = MultiVariable()
        mv.x = Variable(1)
        del mv.x
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            mv.x = Variable(9)
        assert mv.x._value == 9
        assert mv.component_names == ["x"]

    def test_del_plain_attribute_is_normal_python(self):
        mv = MultiVariable()
        mv.note = "memo"
        del mv.note
        assert not hasattr(mv, "note")
        with pytest.raises(AttributeError):
            del mv.missing


class TestMultiVariableClassNodeKwargs:
    """A template's constructor takes the same node keywords as
    ``MultiVariable`` (``description=`` …), and ``compute()`` parameters
    are filled only from what the caller passed or from the declared
    defaults — never from a node attribute that happens to share the
    name."""

    def test_description_kwarg_describes_the_node(self):
        class Cohort(MultiVariableClass):
            def compute(self, start=0):
                self.users = Variable(start)

        c = Cohort(start=1, description="d")
        assert c.description == "d"
        assert c.users._value == 1

    def test_description_also_reaches_compute_when_declared(self):
        seen = {}

        class Cohort(MultiVariableClass):
            def compute(self, start=0, description=None):
                seen["description"] = description

        c = Cohort(start=1, description="d")
        assert seen == {"description": "d"}
        assert c.description == "d"

    def test_params_named_like_read_only_properties_reach_compute(self):
        seen = {}

        class Named(MultiVariableClass):
            def compute(self, code=7, python_name=None, path=None):
                seen.update(code=code, python_name=python_name, path=path)

        n = Named(code=99, python_name="p", path="x/y")
        assert seen == {"code": 99, "python_name": "p", "path": "x/y"}
        # the node's own properties are untouched
        assert isinstance(n.code, str)
        assert n.python_name is None

    def test_excel_props_param_reaches_compute_and_styles_node(self):
        seen = {}

        class Named(MultiVariableClass):
            def compute(self, excel_props=8):
                seen["excel_props"] = excel_props

        n = Named(excel_props={"bold": True})
        assert seen == {"excel_props": {"bold": True}}
        assert n.excel_props == {"bold": True}

    def test_variable_param_named_code_keeps_formula_link(self):
        seen = {}

        class Named(MultiVariableClass):
            def compute(self, code=7):
                seen["code"] = code

        v = Variable(99)
        Named(code=v)
        assert seen["code"] is v

    def test_omitted_params_get_declared_defaults(self):
        seen = []

        class Named(MultiVariableClass):
            def compute(self, description="default-desc", code=7, at=3,
                        add=1, remove=2, excel_props=8):
                seen.append([description, code, at, add, remove, excel_props])

        n = Named()
        assert seen == [["default-desc", 7, 3, 1, 2, 8]]
        # the default belongs to compute(), not to the node
        assert n.description is None
        assert n.excel_props == {}
        assert callable(n.add)

    def test_omitted_display_name_gets_declared_default(self):
        seen = {}

        class Named(MultiVariableClass):
            def compute(self, display_name="dflt"):
                seen["display_name"] = display_name

        Named()
        assert seen == {"display_name": "dflt"}

    def test_omitted_required_param_named_like_property_raises(self):
        class Named(MultiVariableClass):
            def compute(self, code):
                self.x = Variable(code)

        with pytest.raises(TypeError, match="code"):
            Named()

    def test_shell_accepts_params_named_like_read_only_properties(self):
        class Named(MultiVariableClass):
            def compute(self, code=7):
                self.x = Variable(code)

        n = Named.__shell__(code=99, description="d")
        assert n.component_names == []
        assert n.description == "d"


class TestAdoptionRefusals:
    """A container assigned under a second name inside its own tree —
    ``model.m = model`` — used to recurse until Python gave up.
    Refused with the alias to write, spelled in the author's own names;
    the container is untouched."""

    def test_a_container_cannot_hold_itself(self):
        m = Model("m", default_grain="year", default_start="2026", default_periods=2)
        with pytest.raises(ValueError, match=r"cannot hold itself.*`with m as alias:`"):
            m.alias = m
        assert "alias" not in m._components
        assert type(m) is Model  # not demoted on the way to the refusal

    def test_a_container_cannot_hold_its_ancestor(self):
        m = Model("m", default_grain="year", default_start="2026", default_periods=2)
        m.sub = MultiVariable("sub")
        with pytest.raises(
            ValueError, match=r"`m\.sub\.back = m`: .*one of its ancestors.*`with m as back:`"
        ):
            m.sub.back = m
        assert "back" not in m.sub._components

    def test_the_hint_spells_the_names_the_author_bound(self):
        """``mo.Model('m')`` bound as ``model``: the hint says ``model``;
        a holder is spelled from the highest bound ancestor
        (``model.sub``, not the block's alias ``s``), an unattached
        container by its own name."""
        model = Model("m", default_grain="year", default_start="2026", default_periods=2)
        with model:
            model.sub = MultiVariable("sub")
            with model.sub as s:
                with pytest.raises(
                    ValueError, match=r"`model\.sub\.back = model`: .*`with model as back:`"
                ):
                    s.back = model
        with MultiVariable("Debt", loop=True) as debt:
            with pytest.raises(ValueError, match=r"`debt\.me = debt`: .*`with debt as me:`"):
                debt.me = debt

    def test_the_same_name_points_at_a_plain_block(self):
        m = Model("m")
        with pytest.raises(ValueError, match=r"`with m:`"):
            m.m = m

    def test_a_missing_component_says_how_to_assign_it(self):
        m = Model("m")
        m.sub = MultiVariable("sub")
        with pytest.raises(
            AttributeError,
            match=r"has no component 'rate'.*`m\.sub\.rate = mo\.Variable\(\.\.\.\)`.*loop=True",
        ):
            m.sub.rate

    def test_the_missing_component_hint_is_python(self):
        """The display name ('Income Statement') and the declared name
        ('m' bound as ``model``) are never the hint's spelling."""
        with MultiVariable("Income Statement") as income:
            with pytest.raises(AttributeError, match=r"`income\.revenue = mo\.Variable"):
                income.revenue
        model = Model("m")
        with pytest.raises(AttributeError, match=r"`model\.revenue = mo\.Variable"):
            model.revenue
        # A probe with a default still answers the default.
        assert getattr(model, "revenue", None) is None

    def test_the_hint_survives_the_line_it_points_at(self):
        """A comprehension's variable is gone before the line the hint
        says to write, so the hint is spelled from the model (``model.b``);
        a global the frame's own parameter shadows is not the model's
        name there; a setting of the container is never offered as a row."""
        model = Model("m", default_grain="year", default_start="2026", default_periods=2)
        model.a = MultiVariable("a")
        model.b = MultiVariable("b")
        model.a.revenue = Variable([1.0, 2.0])
        with pytest.raises(AttributeError, match=r"`model\.b\.revenue = mo\.Variable"):
            [seg.revenue for seg in (model.a, model.b)]
        with pytest.raises(AttributeError, match=r"`model\.b\.revenue = mo\.Variable"):
            (lambda q: q.revenue)(model.b)

        base = model
        other = Model("other")
        other.b = MultiVariable("b")

        def read(model):  # the parameter shadows the name ``model``
            return base.b.revenue

        with pytest.raises(AttributeError, match=r"`base\.b\.revenue = mo\.Variable"):
            read(other)
        with pytest.raises(AttributeError, match=r"has no component 'default_periods'$"):
            Model("bare").default_periods
        with pytest.raises(AttributeError, match=r"has no component 'default_zone'$"):
            Model("bare").default_zone


class TestAHorizonByAssignment:
    @pytest.fixture(autouse=True)
    def _no_model_noted(self):
        # The model last given a horizon is process-wide state: start from
        # none, so a model an earlier test built cannot answer for this one.
        from dataclasses import replace

        from modeleon.core import mv_context

        ctx = mv_context._loop_context
        token = ctx.set(replace(ctx.get(), windowed=None))
        yield
        ctx.reset(token)

    def test_a_model_given_its_horizon_by_assignment_is_the_window_time_finds(self):
        """``m.default_periods = 3`` marks the model last given a horizon,
        as the constructor does: a line reading ``mo.time`` outside the
        model's ``with`` block (its lines written without the block) finds
        the window there (it used to raise "there is no model window here")."""
        import modeleon as mo

        m = Model("m")
        m.default_grain = "quarter"
        m.default_periods = 3
        m.default_start = "2031-Q2"
        m.pnl = MultiVariable("pnl")
        m.pnl.days = mo.time.days
        assert list(m.pnl.days.value) == [91, 92, 92]

    def test_a_refused_change_of_horizon_notes_nothing(self):
        """A model whose lines already read the window refuses a new
        horizon; the refused model is not the one ``mo.time`` finds next."""
        import modeleon as mo

        m = Model("m", default_grain="month", default_start="2026-01", default_periods=3)
        m.pnl = MultiVariable("pnl")
        m.pnl.days = mo.time.days
        other = Model("other", default_grain="month", default_start="2031-05", default_periods=2)
        with pytest.raises(ValueError):
            m.default_periods = 99
        x = MultiVariable("x")
        x.d = mo.time.days
        assert list(x.d.value) == [31, 30]      # other's window, not m's


class TestAPlainValueUnderAComponentName:
    """A line overwritten by a plain number keeps its name: removing it,
    deleting it or moving its container works as for a line."""

    def _model(self):
        m = Model("m", default_grain="month", default_start="2026-01", default_periods=4)
        with m:
            m.inputs = MultiVariable("Inputs")
            m.inputs.rate = Variable(0.05, display_name="rate")
            m.inputs.a = Variable([1.0, 2.0, 3.0, 4.0], display_name="a")
            m.y = MultiVariable("Y", default_grain="month", default_start="2026-03",
                                   default_periods=4)
        m.inputs.rate = 0.05
        return m

    def test_remove(self):
        m = self._model()
        assert m.inputs.remove("rate") == 0.05
        assert "rate" not in m.inputs._component_names

    def test_del(self):
        m = self._model()
        del m.inputs.rate
        assert "rate" not in m.inputs._component_names

    def test_its_container_moved(self):
        m = self._model()
        inputs = m.remove("inputs")
        m.y.inputs = inputs
        assert m.y.inputs.a.time.start == "2026-03"
        assert m.y.inputs.a.id == "m.y.inputs.a"
