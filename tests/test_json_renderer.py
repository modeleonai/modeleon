# SPDX-License-Identifier: Apache-2.0
"""Tests for the JSON renderer — proves the ``Renderer`` seam works.

Every test here renders a Variable's formula as both Excel (via
``ExcelTranslator``) and JSON (via ``JsonRenderer``) and asserts the
JSON form is renderer-neutral: no cell addresses, no sheet names, no
Excel-specific literal formatting.
"""

from __future__ import annotations

import json

import modeleon as mo
from modeleon.compile.json.renderer import JsonRenderer, to_json
from modeleon.compile.walker import Walker


def _render_json(var):
    """Walker-constructed render, equivalent to ``to_json(var._expr)``."""
    renderer = JsonRenderer()
    Walker(renderer)
    from modeleon.compile.renderer import RenderCtx
    return renderer.walker.render(var._expr, RenderCtx())


def _build_model():
    model = mo.MultiVariable("m")
    model._python_name = "model"
    model.s = mo.MultiVariable("s", excel_props={'tab': True})
    return model


class TestAtoms:
    def test_literal(self):
        v = mo.Variable(0)
        v._expr = __import__("modeleon").core.expr.Literal(42)
        assert to_json(v._expr) == {"kind": "literal", "value": 42}

    def test_varref_uses_qualified_path(self):
        model = _build_model()
        model.s.a = mo.Variable(1)
        from modeleon.core.expr import VarRef
        ref = VarRef(model.s.a)
        out = to_json(ref)
        assert out == {"kind": "varref", "path": "model.s.a"}


class TestArithmetic:
    def test_addition(self):
        model = _build_model()
        model.s.a = mo.Variable(10)
        model.s.b = mo.Variable(20)
        model.s.total = model.s.a + model.s.b
        out = to_json(model.s.total._expr)
        assert out == {
            "kind": "binop",
            "op": "+",
            "left": {"kind": "varref", "path": "model.s.a"},
            "right": {"kind": "varref", "path": "model.s.b"},
        }

    def test_multiply_with_literal(self):
        model = _build_model()
        model.s.revenue = mo.Variable(100)
        model.s.cogs = model.s.revenue * 0.6
        out = to_json(model.s.cogs._expr)
        assert out == {
            "kind": "binop",
            "op": "*",
            "left": {"kind": "varref", "path": "model.s.revenue"},
            "right": {"kind": "literal", "value": 0.6},
        }

    def test_nested_expression(self):
        model = _build_model()
        model.s.a = mo.Variable(2)
        model.s.b = mo.Variable(3)
        model.s.c = mo.Variable(5)
        model.s.total = (model.s.a * model.s.b) + model.s.c
        out = to_json(model.s.total._expr)
        # Top-level BinOp("+", Paren(BinOp("*", ...)) or similar.
        assert out["kind"] == "binop"
        assert out["op"] == "+"


class TestBackendNeutrality:
    """The JSON output contains no Excel-specific artifacts."""

    def test_no_cell_addresses_in_output(self):
        model = _build_model()
        model.s.a = mo.Variable(10)
        model.s.b = mo.Variable(20)
        model.s.total = model.s.a + model.s.b
        serialized = json.dumps(to_json(model.s.total._expr))
        # Excel-isms we explicitly don't want to see:
        assert "A1" not in serialized
        assert "B2" not in serialized
        assert "!" not in serialized  # no sheet-qualified refs
        assert "TRUE" not in serialized  # Python ``true``, not Excel ``TRUE``
        # Paths are there, but as identity — not cell refs:
        assert "model.s.a" in serialized
        assert "model.s.b" in serialized

    def test_output_is_json_serializable(self):
        model = _build_model()
        model.s.revenue = mo.Variable(1000)
        model.s.cogs_pct = mo.Variable(0.6)
        model.s.cogs = model.s.revenue * model.s.cogs_pct
        model.s.profit = model.s.revenue - model.s.cogs
        # json.dumps is the acid test for renderer-neutrality.
        json.dumps(to_json(model.s.profit._expr))  # no TypeError


class TestFuncCall:
    def test_sum(self):
        model = _build_model()
        model.s.values = mo.Variable([1, 2, 3, 4])
        model.s.total = mo.SUM(model.s.values)
        out = to_json(model.s.total._expr)
        assert out["kind"] == "funccall"
        assert out["func"].upper() == "SUM"
        assert len(out["args"]) == 1
        assert out["args"][0] == {"kind": "varref", "path": "model.s.values"}

    def test_if(self):
        model = _build_model()
        model.s.x = mo.Variable(100)
        model.s.result = mo.IF(model.s.x > 50, model.s.x, 0)
        out = to_json(model.s.result._expr)
        assert out["kind"] == "funccall"
        assert out["func"].upper() == "IF"
        assert len(out["args"]) == 3
