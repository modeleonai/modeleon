# SPDX-License-Identifier: Apache-2.0
"""Public API surface — stabilized extension points.

- ``.value``, ``.expr``, ``.formula``, ``.source_code`` are public
  properties. Underscored forms (``_value``, ``_expr``,
  ``_source_code``) remain internal storage and may change.
- ``modeleon.extend`` is the stable surface for writing custom DSL
  functions — re-exports the AST node types, the builder helpers,
  and the ``@pyformula`` decorator.
"""

from __future__ import annotations

import modeleon as mo


class TestPublicProperties:
    """The four public Variable properties — read + write where applicable."""

    def test_value_readable(self):
        v = mo.Variable(42)
        assert v.value == 42

    def test_expr_readable(self):
        v = mo.Variable(1) + mo.Variable(2)
        # ``.expr`` exposes the AST root — a BinOp for this expression.
        assert v.expr is not None
        assert v.expr.op == "+"

    def test_formula_readable(self):
        a = mo.Variable(1)
        b = mo.Variable(2)
        total = a + b
        # ``.formula`` renders from AST — symbolic names (if present) or
        # fallback identifiers.
        assert total.formula is not None
        assert "+" in total.formula

    def test_source_code_captures_resolved_call(self):
        """Financial helpers stamp ``.source_code`` with resolved scalar
        values — distinct from the AST-derived ``.formula``."""
        r = mo.PMT(0.005, 360, 300_000)
        # ``.source_code`` captures the concrete call.
        assert r.source_code == "PMT(0.005, 360, 300000, 0, 0)"
        # ``.formula`` is derived from the AST — same funcname, but
        # literal node rendering. It's fine if the two differ; the
        # contract is just that source_code is non-None when the helper
        # set it.
        assert r.formula is not None

    def test_source_code_none_for_plain_input(self):
        v = mo.Variable(42)
        assert v.source_code is None

    def test_source_code_writable(self):
        v = mo.Variable(1) + mo.Variable(2)
        v.source_code = "my_custom(1, 2)"
        assert v.source_code == "my_custom(1, 2)"
        v.source_code = None
        assert v.source_code is None


class TestExtendModule:
    """``modeleon.extend`` re-exports the extension-author surface."""

    def test_pyformula_importable(self):
        from modeleon.extend import pyformula

        @pyformula
        def double(x):
            return x * 2

        r = double(mo.Variable(5))
        assert r.value == 10
        assert r.expr.func == "DOUBLE"

    def test_make_func_var_importable(self):
        from modeleon.extend import FuncCall, VarRef, make_func_var

        v = mo.Variable(3.0)
        r = make_func_var(
            "TRIPLE", [VarRef(v)], v.value * 3, value_type="float",
        )
        assert r.value == 9.0
        assert isinstance(r.expr, FuncCall)
        assert r.expr.func == "TRIPLE"

    def test_ast_nodes_importable(self):
        from modeleon.extend import (
            BinOp,
            Expr,
            FuncCall,
            Literal,
            VarRef,
        )
        # Just checking the symbols exist and inherit from Expr where
        # expected — downstream renderers do isinstance checks.
        assert issubclass(BinOp, Expr)
        assert issubclass(FuncCall, Expr)
        assert issubclass(Literal, Expr)
        assert issubclass(VarRef, Expr)

    def test_render_ctx_importable(self):
        from modeleon.extend import RenderCtx
        ctx = RenderCtx()
        # Default period_idx is 0 — stable contract for callers.
        assert ctx.period_idx == 0
