# SPDX-License-Identifier: Apache-2.0
"""Renderer protocol for AST rendering.

An ``Expr`` tree (see :mod:`modeleon.core.expr`) is renderer-neutral —
it describes *what* a formula computes, not *how* it renders. A
:class:`Renderer` implementation decides the rendering: Excel formula
strings, JSON dicts, pandas expressions, SQL, etc.

The :class:`Walker` (see :mod:`.walker`) dispatches each AST node type
to the matching ``render_*`` method on a renderer. Backends use the
walker to recurse into child nodes.

This file defines the Protocol (structural typing) that every renderer
must satisfy, plus the per-call context object.
"""

from __future__ import annotations

import warnings
from dataclasses import dataclass
from typing import Any, Optional, Protocol, TYPE_CHECKING

if TYPE_CHECKING:
    from ..core.expr import (
        BinOp,
        Compare,
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


@dataclass
class RenderCtx:
    """Per-call state threaded through the AST walk.

    Carried from the top-level ``translate``/``render`` call down into
    every recursive ``Walker.render`` invocation. Backends read from
    it; most won't mutate. The one exception is
    ``SelfRef``-period-0 rendering, which may flip a flag so the inner
    ``{prev}`` slot renders as the start value rather than the previous
    period's cell reference.
    """

    period_idx: int = 0
    current_sheet: Optional[str] = None
    # Renderer-specific address payload for the Variable currently being
    # rendered (e.g. Excel uses it to resolve SelfRef's ``{prev}`` to the
    # previous period's cell). Typed as ``Any`` so the generic protocol
    # stays renderer-neutral; each concrete renderer narrows as needed.
    self_address: Optional[Any] = None
    # The Variable whose ``_expr`` is the top-level tree being rendered.
    # Available to fallback paths so they can inline ``self_var.value``
    # when an AST node (typically a backend-specific ``FuncCall``) can't
    # be rendered natively. Callers set this at the top-level
    # ``translate()`` / ``render()`` call; deeper recursion inherits.
    self_var: Optional[Any] = None


class Renderer(Protocol):
    """Structural protocol for AST renderers.

    Every AST node type has a corresponding ``render_*`` method. The
    :class:`Walker` dispatches ``type(node) -> renderer.render_<name>``
    and passes the node + context. Renderers recurse by calling back
    into the walker (``self.walker.render(child, ctx)``).

    Return type is ``Any`` at the protocol level — concrete renderers
    narrow it (``str`` for Excel, ``dict`` for JSON, etc.).

    **Scope: rendering only.** Renderers translate an AST node into
    their target representation; they never compute values. Value
    computation is Python-eager and lives in ``core/variable_ops.py``
    (single chokepoint).

    **Fallback contract.** Renderers MUST gracefully degrade on AST
    nodes they cannot natively render (unknown ``FuncCall``, out-of-
    scope ``VarRef``, etc.): inline ``var._value`` as a literal in the
    target format, or emit a placeholder + warning when the value is
    ``None``. This keeps backend-specific functions (pandas-native
    ``EWMA``, SQL-only ``LAG``, etc.) from breaking other renderers —
    they render natively in their home backend and value-inline
    elsewhere. Helper: :func:`fallback_to_value` below.
    """

    walker: "Walker"

    def render_literal(self, node: "Literal", ctx: RenderCtx) -> Any: ...
    def render_varref(self, node: "VarRef", ctx: RenderCtx) -> Any: ...
    def render_paren(self, node: "Paren", ctx: RenderCtx) -> Any: ...
    def render_binop(self, node: "BinOp", ctx: RenderCtx) -> Any: ...
    def render_unaryop(self, node: "UnaryOp", ctx: RenderCtx) -> Any: ...
    def render_compare(self, node: "Compare", ctx: RenderCtx) -> Any: ...
    def render_subscript(self, node: "Subscript", ctx: RenderCtx) -> Any: ...
    def render_listexpr(self, node: "ListExpr", ctx: RenderCtx) -> Any: ...
    def render_methodcall(self, node: "MethodCall", ctx: RenderCtx) -> Any: ...
    def render_funccall(self, node: "FuncCall", ctx: RenderCtx) -> Any: ...
    def render_selfref(self, node: "SelfRef", ctx: RenderCtx) -> Any: ...
    def render_rollingaggregate(
        self, node: "RollingAggregate", ctx: RenderCtx
    ) -> Any: ...
    def render_regrain(
        self, node: "Regrain", ctx: RenderCtx
    ) -> Any: ...


_PLACEHOLDER = object()
"""Sentinel returned by :func:`fallback_to_value` when a renderer can't
natively handle a node AND no ``_value`` is available. Concrete
renderers decide how to format it (blank cell, ``null``, etc.)."""


def fallback_to_value(
    var: Any,
    *,
    renderer_name: str,
    node_description: str = "this node",
) -> Any:
    """Return ``var.value`` if it's available, else the placeholder
    sentinel (plus a warning).

    Called by renderers when they encounter an AST node they can't
    render natively — e.g. a :class:`FuncCall` whose ``render_backends``
    doesn't include this renderer, or a :class:`VarRef` to a Variable
    outside the current emission scope. The renderer receives the
    Python value (or :data:`_PLACEHOLDER`) and is responsible for
    formatting it into its target (Excel literal, JSON dict, SQL
    literal, etc.).
    """
    value = getattr(var, "value", None)
    if value is not None:
        return value
    warnings.warn(
        f"{renderer_name}: no native render for {node_description} "
        f"and no Python-computed value available; emitting placeholder.",
        stacklevel=3,
    )
    return _PLACEHOLDER


# Forward reference to avoid a circular import with walker.py.
from .walker import Walker  # noqa: E402
