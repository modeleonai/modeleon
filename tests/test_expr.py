# SPDX-License-Identifier: Apache-2.0
"""Tests for the expression-tree nodes in modeleon.core.expr.

These tests focus on two properties:

1. ``to_string()`` must reproduce the exact formula-string form the DSL
   produces today. This lets us migrate ``Variable._expr`` storage
   incrementally without breaking JSON exports, downstream consumers,
   or the existing regex-based translator.

2. ``iter_refs()`` must yield every dependency, so ``_dependency_refs`` can
   eventually be derived from the tree instead of maintained manually.
"""

import pytest

import modeleon as mo
from modeleon.core.expr import (
    BinOp,
    Compare,
    Expr,
    FuncCall,
    ListExpr,
    Literal,
    MethodCall,
    SelfRef,
    Subscript,
    UnaryOp,
    VarRef,
)


# ─── Atoms ──────────────────────────────────────────────────────


class TestLiteral:
    def test_int_renders_as_string(self):
        assert Literal(42).to_string() == "42"

    def test_float_renders_as_string(self):
        assert Literal(0.65).to_string() == "0.65"

    def test_string_renders_python_quoted(self):
        # Strings in a formula string render Python-style quoted so
        # ``Variable.formula`` round-trips cleanly. The translator emits
        # the same Literal with Excel-style double quotes for the cell.
        assert Literal("Bear").to_string() == "'Bear'"

    def test_literal_has_no_refs(self):
        assert list(Literal(1).iter_refs()) == []


class TestVarRef:
    def test_renders_python_name(self):
        v = mo.Variable(100)
        v.set_python_name("revenue")
        assert VarRef(v).to_string() == "revenue"

    def test_falls_back_to_path_leaf(self):
        # Stash the Variable inside a list inside an inner function so
        # it's not bound to a local name at any scope lazy-discovery can
        # walk to. (Avoids pytest's assertion rewriting injecting
        # @py_assert* locals that discovery would match.)
        def render():
            vs = [mo.Variable(100)]
            return VarRef(vs[0]).to_string(), vs[0].path.leaf
        rendered, leaf = render()
        # Formula rendering uses leaf — short identifier, not the full
        # dotted path, for readable Python / Excel formula strings.
        assert rendered == leaf

    def test_iter_refs_yields_self(self):
        v = mo.Variable(100)
        assert list(VarRef(v).iter_refs()) == [v]

    def test_explicit_set_python_name_takes_precedence(self):
        """Explicitly setting ``python_name`` makes ``VarRef.to_string``
        render with the new name, overriding whatever lazy discovery
        may have picked up."""
        def build():
            vs = [mo.Variable(100)]
            vs[0].set_python_name("revenue")
            return VarRef(vs[0]).to_string()
        assert build() == "revenue"


# ─── Arithmetic ────────────────────────────────────────────────


class TestBinOp:
    def test_renders_infix(self):
        a = mo.Variable(10)
        a.set_python_name("a")
        b = mo.Variable(20)
        b.set_python_name("b")
        expr = BinOp("+", VarRef(a), VarRef(b))
        assert expr.to_string() == "a + b"

    def test_scalar_right_operand(self):
        a = mo.Variable(10)
        a.set_python_name("revenue")
        expr = BinOp("*", VarRef(a), Literal(0.65))
        assert expr.to_string() == "revenue * 0.65"

    def test_matches_dsl_output(self):
        """Round-trip: a + b constructed via the DSL should match the AST form."""
        a = mo.Variable(10, display_name="A")
        a.set_python_name("a")
        b = mo.Variable(20, display_name="B")
        b.set_python_name("b")
        c = a + b
        manual = BinOp("+", VarRef(a), VarRef(b))
        assert c.formula == manual.to_string()

    def test_iter_refs_yields_both_sides(self):
        a = mo.Variable(1)
        b = mo.Variable(2)
        expr = BinOp("*", VarRef(a), VarRef(b))
        assert list(expr.iter_refs()) == [a, b]


class TestUnaryOp:
    def test_negation(self):
        v = mo.Variable(10)
        v.set_python_name("revenue")
        assert UnaryOp("-", VarRef(v)).to_string() == "-revenue"

    def test_iter_refs(self):
        v = mo.Variable(10)
        assert list(UnaryOp("-", VarRef(v)).iter_refs()) == [v]


class TestCompare:
    def test_eq(self):
        a = mo.Variable(10)
        a.set_python_name("a")
        b = mo.Variable(20)
        b.set_python_name("b")
        assert Compare("==", VarRef(a), VarRef(b)).to_string() == "a == b"

    def test_all_operators(self):
        v = mo.Variable(1)
        v.set_python_name("x")
        for op in ("==", "!=", "<", "<=", ">", ">="):
            assert Compare(op, VarRef(v), Literal(0)).to_string() == f"x {op} 0"


# ─── Subscript / list ────────────────────────────────────────


class TestSubscript:
    def test_positional_index(self):
        v = mo.Variable([1, 2, 3])
        v.set_python_name("revenue")
        assert Subscript(VarRef(v), 0).to_string() == "revenue[0]"

    def test_named_key(self):
        v = mo.Variable({"Bear": 1, "Base": 2, "Bull": 3})
        v.set_python_name("scenarios")
        assert Subscript(VarRef(v), "Bear").to_string() == 'scenarios["Bear"]'

    def test_slice(self):
        v = mo.Variable([1, 2, 3, 4, 5])
        v.set_python_name("revenue")
        assert Subscript(VarRef(v), slice(1, 4)).to_string() == "revenue[1:4]"

    def test_slice_with_step(self):
        v = mo.Variable([1, 2, 3, 4])
        v.set_python_name("x")
        assert Subscript(VarRef(v), slice(0, 4, 2)).to_string() == "x[0:4:2]"

    def test_iter_refs_follows_base(self):
        v = mo.Variable(1)
        assert list(Subscript(VarRef(v), 0).iter_refs()) == [v]


class TestListExpr:
    def test_empty(self):
        assert ListExpr([]).to_string() == "[]"

    def test_mixed_vars_and_literals(self):
        a = mo.Variable(1)
        a.set_python_name("revenue")
        b = mo.Variable(2)
        b.set_python_name("costs")
        expr = ListExpr([VarRef(a), Literal(0), VarRef(b), Literal(5)])
        assert expr.to_string() == "[revenue, 0, costs, 5]"

    def test_iter_refs_only_vars(self):
        a = mo.Variable(1)
        b = mo.Variable(2)
        expr = ListExpr([VarRef(a), Literal(100), VarRef(b)])
        assert list(expr.iter_refs()) == [a, b]


# ─── Function / method calls ────────────────────────────────


class TestFuncCall:
    def test_no_args(self):
        assert FuncCall("today", []).to_string() == "today()"

    def test_if(self):
        cond = mo.Variable(1)
        cond.set_python_name("revenue")
        thresh = Literal(1000)
        expr = FuncCall("IF", [Compare(">", VarRef(cond), thresh), Literal(100), Literal(0)])
        assert expr.to_string() == "IF(revenue > 1000, 100, 0)"

    def test_sum_single_arg(self):
        v = mo.Variable([1, 2, 3])
        v.set_python_name("values")
        assert FuncCall("sum", [VarRef(v)]).to_string() == "sum(values)"

    def test_max_multi_arg(self):
        a = mo.Variable(1); a.set_python_name("a")
        b = mo.Variable(2); b.set_python_name("b")
        c = mo.Variable(3); c.set_python_name("c")
        assert FuncCall("max", [VarRef(a), VarRef(b), VarRef(c)]).to_string() == "max(a, b, c)"

    def test_iter_refs_follows_all_args(self):
        a = mo.Variable(1)
        b = mo.Variable(2)
        expr = FuncCall("IF", [Compare(">", VarRef(a), Literal(0)), VarRef(b), Literal(0)])
        assert list(expr.iter_refs()) == [a, b]


class TestMethodCall:
    def test_no_args(self):
        v = mo.Variable(1)
        v.set_python_name("revenue")
        assert MethodCall(VarRef(v), "copy", [], {}).to_string() == "revenue.copy()"

    def test_positional_arg(self):
        v = mo.Variable(1)
        v.set_python_name("revenue")
        assert MethodCall(VarRef(v), "shift", [1], {}).to_string() == "revenue.shift(1)"

    def test_positional_and_kwarg(self):
        v = mo.Variable(1)
        v.set_python_name("debt")
        assert (
            MethodCall(VarRef(v), "shift", [1], {"fill_value": 0}).to_string()
            == "debt.shift(1, fill_value=0)"
        )

    def test_kwarg_only(self):
        v = mo.Variable(1)
        v.set_python_name("revenue")
        assert (
            MethodCall(VarRef(v), "resample", [], {"method": "sum"}).to_string()
            == "revenue.resample(method=sum)"
        )

    def test_iter_refs_follows_base_and_expr_args(self):
        a = mo.Variable(1)
        b = mo.Variable(2)
        expr = MethodCall(VarRef(a), "something", [VarRef(b)], {})
        assert list(expr.iter_refs()) == [a, b]


# ─── Self-reference ──────────────────────────────────────────


class TestSelfRef:
    def test_basic(self):
        flow = mo.Variable(100)
        flow.set_python_name("cash_flow")
        expr = SelfRef(
            start=Literal(0),
            template="{prev} + {flow}",
            variables={"flow": VarRef(flow)},
            periods=5,
        )
        assert expr.to_string() == 'recurrence(0, "{prev} + {flow}", flow=cash_flow, periods=5)'

    def test_variable_start(self):
        opening = mo.Variable(1000)
        opening.set_python_name("opening_cash")
        flow = mo.Variable(100)
        flow.set_python_name("flow")
        expr = SelfRef(
            start=VarRef(opening),
            template="{prev} + {flow}",
            variables={"flow": VarRef(flow)},
            periods=None,
        )
        assert expr.to_string() == 'recurrence(opening_cash, "{prev} + {flow}", flow=flow)'

    def test_multiple_template_variables(self):
        churn = mo.Variable(0.05)
        churn.set_python_name("churn")
        new = mo.Variable(10)
        new.set_python_name("new")
        expr = SelfRef(
            start=Literal(1000),
            template="{prev} * (1 - {churn}) + {new}",
            variables={"churn": VarRef(churn), "new": VarRef(new)},
            periods=12,
        )
        rendered = expr.to_string()
        assert rendered.startswith('recurrence(1000, "{prev} * (1 - {churn}) + {new}"')
        assert "churn=churn" in rendered
        assert "new=new" in rendered
        assert rendered.endswith("periods=12)")

    def test_iter_refs_collects_start_and_template_vars(self):
        start_v = mo.Variable(10)
        flow_v = mo.Variable(20)
        expr = SelfRef(
            start=VarRef(start_v),
            template="{prev} + {flow}",
            variables={"flow": VarRef(flow_v)},
            periods=None,
        )
        assert list(expr.iter_refs()) == [start_v, flow_v]


# ─── Full expression shapes (integration-ish) ───────────────


class TestComplexExpressions:
    def test_nested_arithmetic(self):
        """Mirrors the DSL's output for (a + b) * c without precedence parens
        (engine's binary op emits the flat form today)."""
        a = mo.Variable(1); a.set_python_name("a")
        b = mo.Variable(2); b.set_python_name("b")
        c = mo.Variable(3); c.set_python_name("c")
        # In the AST, we'd build: BinOp("*", BinOp("+", a, b), c)
        expr = BinOp("*", BinOp("+", VarRef(a), VarRef(b)), VarRef(c))
        assert expr.to_string() == "a + b * c"

    def test_iter_refs_dedup_not_guaranteed(self):
        """Spec: iter_refs is allowed to yield duplicates for shared subtrees.
        Callers that want uniqueness should dedupe themselves."""
        a = mo.Variable(1)
        expr = BinOp("+", VarRef(a), VarRef(a))
        refs = list(expr.iter_refs())
        assert refs == [a, a]  # two yields, same object

    def test_full_pnl_shape(self):
        """Shape that mirrors a real P&L formula chain."""
        revenue = mo.Variable(1_000_000); revenue.set_python_name("revenue")
        cogs_pct = mo.Variable(0.6); cogs_pct.set_python_name("cogs_pct")

        cogs_expr = BinOp("*", VarRef(revenue), VarRef(cogs_pct))
        gross_expr = BinOp("-", VarRef(revenue), cogs_expr)

        assert cogs_expr.to_string() == "revenue * cogs_pct"
        assert gross_expr.to_string() == "revenue - revenue * cogs_pct"

        assert list(gross_expr.iter_refs()) == [revenue, revenue, cogs_pct]


# ─── Base class contract ────────────────────────────────────


class TestBaseContract:
    def test_base_class_methods_raise(self):
        """Ensure the abstract-ish base correctly forces subclass overrides."""
        with pytest.raises(NotImplementedError):
            Expr().to_string()
        with pytest.raises(NotImplementedError):
            list(Expr().iter_refs())
