# SPDX-License-Identifier: Apache-2.0
"""Public extension surface for writing custom DSL functions.

Most users never need this module — arithmetic on :class:`Variable`
builds the AST automatically, and the shipped helpers in
:mod:`modeleon.functions` (``SUM``, ``IRR``, ``IF``, …) cover the
common cases. Reach for :mod:`modeleon.extend` when you want to add a
named function of your own — one that renders as ``=YOURNAME(args)``
in Excel, shows up in the graph with tracked deps, and round-trips
through :func:`modeleon.to_json` like a built-in.

Two entry points, ordered by how much control you want:

- :func:`pyformula` — decorator. Wraps an arbitrary Python function so
  calling it with Variable args returns a Variable whose AST is a
  :class:`FuncCall`. Deps are preserved; every renderer falls back to
  the computed value. Use when the body is opaque Python
  (``scipy.optimize``, a trained model, an HTTP call).

- :func:`make_func_var` — the raw builder. Use when you need explicit
  control over ``render_backends`` / ``compute_backends``, a custom
  ``source_code`` string, or non-trivial arg handling.

The AST node classes (``Expr``, ``FuncCall``, ``VarRef``,
``Literal``, :class:`RenderCtx`) are re-exported so custom renderers
and graph walkers can type-check against the same surface the engine
uses internally.

Stability: this is the public extension API. Semver applies at the
minor level — we won't break it in patches, and breaking changes in
minor releases get a CHANGELOG note.
"""

from __future__ import annotations

from .compile.renderer import RenderCtx
from .core.expr import (
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
from .functions._helpers import (
    make_func_var,
    operand,
    pyformula,
    src,
    text_operand,
    val,
)


__all__ = [
    # Decorator for opaque-Python helpers.
    "pyformula",
    # Lower-level builder + arg-splitting helpers.
    "make_func_var",
    "operand",
    "text_operand",
    "src",
    "val",
    # AST node classes (for custom renderers / graph walkers).
    "Expr",
    "Literal",
    "VarRef",
    "Paren",
    "BinOp",
    "UnaryOp",
    "Compare",
    "Subscript",
    "ListExpr",
    "FuncCall",
    "MethodCall",
    "SelfRef",
    "RollingAggregate",
    # Renderer context for walking an AST.
    "RenderCtx",
]
