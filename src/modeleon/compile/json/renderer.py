# SPDX-License-Identifier: Apache-2.0
"""JSON renderer — renders an ``Expr`` tree to renderer-neutral dicts.

Exists primarily to prove the :class:`Renderer` seam: writing a second
renderer is cheap and doesn't require touching the core AST, the walker,
or any Excel-specific code. Anyone building a web grid, diff view, or
LLM-consumable formula format can consume these dicts directly.

Example output for ``a + b * 2``::

    {"kind": "binop", "op": "+",
     "left":  {"kind": "varref", "path": "pnl.a"},
     "right": {"kind": "binop", "op": "*",
               "left":  {"kind": "varref", "path": "pnl.b"},
               "right": {"kind": "literal", "value": 2}}}

Nothing here knows about Excel; cell addresses, sheet names, and
Excel-specific literal formatting are absent by construction.
"""

from __future__ import annotations

from typing import Any, Dict, List

from ...core.expr import (
    BinOp,
    Compare,
    Expr,
    FuncCall,
    ListExpr,
    Literal,
    MethodCall,
    Paren,
    RollingAggregate,
    SelfRef,
    Subscript,
    UnaryOp,
    VarRef,
)
from ..renderer import RenderCtx
from ..walker import Walker


class JsonRenderer:
    """Renders AST nodes to JSON-serializable dicts.

    No addresses, no sheet maps — identity is carried by each
    VarRef's ``var.id`` (the qualified path). Consumers that want to
    resolve a VarRef can look up ``path`` in their own map.
    """

    def __init__(self) -> None:
        self.walker: Any = None  # wired by Walker(renderer) constructor

    # ─── Atoms ──────────────────────────────────────────────────

    def render_literal(self, node: Literal, ctx: RenderCtx) -> Dict[str, Any]:
        return {"kind": "literal", "value": node.value}

    def render_varref(self, node: VarRef, ctx: RenderCtx) -> Dict[str, Any]:
        return {"kind": "varref", "path": node.var.id}

    # ─── Structural ────────────────────────────────────────────

    def render_paren(self, node: Paren, ctx: RenderCtx) -> Dict[str, Any]:
        return {"kind": "paren", "inner": self.walker.render(node.inner, ctx)}

    def render_listexpr(self, node: ListExpr, ctx: RenderCtx) -> Dict[str, Any]:
        return {
            "kind": "list",
            "items": [self.walker.render(item, ctx) for item in node.items],
        }

    # ─── Operators ─────────────────────────────────────────────

    def render_binop(self, node: BinOp, ctx: RenderCtx) -> Dict[str, Any]:
        return {
            "kind": "binop",
            "op": node.op,
            "left": self.walker.render(node.left, ctx),
            "right": self.walker.render(node.right, ctx),
        }

    def render_unaryop(self, node: UnaryOp, ctx: RenderCtx) -> Dict[str, Any]:
        return {
            "kind": "unaryop",
            "op": node.op,
            "operand": self.walker.render(node.operand, ctx),
        }

    def render_compare(self, node: Compare, ctx: RenderCtx) -> Dict[str, Any]:
        return {
            "kind": "compare",
            "op": node.op,
            "left": self.walker.render(node.left, ctx),
            "right": self.walker.render(node.right, ctx),
        }

    # ─── Subscript / calls ────────────────────────────────────

    def render_subscript(self, node: Subscript, ctx: RenderCtx) -> Dict[str, Any]:
        key: Any = node.key
        if isinstance(key, slice):
            key = {"kind": "slice", "start": key.start, "stop": key.stop, "step": key.step}
        return {
            "kind": "subscript",
            "base": self.walker.render(node.base, ctx),
            "key": key,
        }

    def render_methodcall(self, node: MethodCall, ctx: RenderCtx) -> Dict[str, Any]:
        args: List[Any] = [
            self.walker.render(a, ctx) if isinstance(a, Expr) else a
            for a in node.args
        ]
        kwargs: Dict[str, Any] = {
            k: (self.walker.render(v, ctx) if isinstance(v, Expr) else v)
            for k, v in node.kwargs.items()
        }
        return {
            "kind": "method",
            "base": self.walker.render(node.base, ctx),
            "method": node.method,
            "args": args,
            "kwargs": kwargs,
        }

    def render_funccall(self, node: FuncCall, ctx: RenderCtx) -> Dict[str, Any]:
        out: Dict[str, Any] = {
            "kind": "funccall",
            "func": node.func,
            "args": [self.walker.render(a, ctx) for a in node.args],
        }
        # Surface backend-specific hints so downstream tooling (LLM
        # formula explainers, graph views, diff tools) knows this
        # function is native only to certain targets. Emit each field
        # only when non-``None`` to keep universal functions terse.
        if node.render_backends is not None:
            out["render_backends"] = sorted(node.render_backends)
        if node.compute_backends is not None:
            out["compute_backends"] = sorted(node.compute_backends)
        return out

    # ─── Time patterns ─────────────────────────────────────────

    def render_selfref(self, node: SelfRef, ctx: RenderCtx) -> Dict[str, Any]:
        return {
            "kind": "selfref",
            "start": self.walker.render(node.start, ctx),
            "template": node.template,
            "variables": {
                name: self.walker.render(expr, ctx)
                for name, expr in node.variables.items()
            },
        }

    def render_rollingaggregate(
        self, node: RollingAggregate, ctx: RenderCtx
    ) -> Dict[str, Any]:
        return {
            "kind": "rolling",
            "source": self.walker.render(node.source, ctx),
            "func": node.func,
            "window": node.window,
            "fill": node.fill,
        }


def to_json(expr: Expr) -> Any:
    """Convenience one-shot: render an ``Expr`` tree as a JSON-ready dict.

    Mirrors :meth:`ExcelTranslator.translate` but for the JSON
    renderer. Useful for diff views, AI / LLM formula inspection, and
    web-grid rendering.
    """
    renderer = JsonRenderer()
    walker = Walker(renderer)
    return walker.render(expr, RenderCtx())
