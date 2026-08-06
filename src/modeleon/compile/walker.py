# SPDX-License-Identifier: Apache-2.0
"""AST walker — dispatches ``Expr`` nodes to a :class:`Renderer`.

Owns the single ``type(node) -> method`` dispatch table shared by
every renderer. Backends provide the rendering; the walker provides the
traversal. Backends recurse by calling ``self.walker.render(child, ctx)``.
"""

from __future__ import annotations

from typing import Any, Dict, TYPE_CHECKING

from ..core.expr import (
    BinOp,
    Compare,
    Expr,
    FuncCall,
    ListExpr,
    Literal,
    MethodCall,
    Paren,
    Regrain,
    RollingAggregate,
    SelfRef,
    Subscript,
    UnaryOp,
    VarRef,
)

if TYPE_CHECKING:
    from .renderer import Renderer, RenderCtx


class Walker:
    """Dispatches AST nodes to the matching ``render_*`` method on a
    renderer. One walker instance per renderer; walker wires itself onto
    the renderer at construction so backends can recurse into children."""

    def __init__(self, renderer: "Renderer") -> None:
        self.renderer = renderer
        renderer.walker = self
        self._dispatch: Dict[type, Any] = {
            Literal:          renderer.render_literal,
            VarRef:           renderer.render_varref,
            Paren:            renderer.render_paren,
            BinOp:            renderer.render_binop,
            UnaryOp:          renderer.render_unaryop,
            Compare:          renderer.render_compare,
            Subscript:        renderer.render_subscript,
            ListExpr:         renderer.render_listexpr,
            MethodCall:       renderer.render_methodcall,
            FuncCall:         renderer.render_funccall,
            SelfRef:          renderer.render_selfref,
            RollingAggregate: renderer.render_rollingaggregate,
            Regrain:          renderer.render_regrain,
        }

    def render(self, node: Expr, ctx: "RenderCtx") -> Any:
        handler = self._dispatch.get(type(node))
        if handler is None:
            raise TypeError(
                f"Renderer has no handler for AST node type "
                f"{type(node).__name__!r}. Add a ``render_{type(node).__name__.lower()}`` "
                f"method to the renderer."
            )
        return handler(node, ctx)
