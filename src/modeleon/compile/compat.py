# SPDX-License-Identifier: Apache-2.0
"""Backend compatibility inspection for an AST, Variable, or MultiVariable.

:func:`check_compat` walks an expression tree (or everything reachable
from a Variable / MultiVariable) and reports every :class:`FuncCall`
whose ``render_backends`` or ``compute_backends`` hint excludes a given
target. An empty list means the tree is fully compatible — a renderer
or compute engine for that target can emit every call natively.

Useful as a pre-flight check before compiling a model to an unfamiliar
target ("will this translate cleanly to pandas?") and as the backing
data for diff / graph views that want to flag fallback-bound nodes.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterator, List, Set

from ..core.expr import Expr, FuncCall, VarRef


@dataclass(frozen=True)
class CompatIssue:
    """A single :class:`FuncCall` that won't render/compute on a target.

    Emitted by :func:`check_compat` for each call whose backend hint
    excludes the requested target. The caller can render these into a
    warning, a UI badge, or a CI assertion.
    """

    func: str
    kind: str                 # "render" or "compute"
    declared: frozenset       # the FuncCall's declared backend set
    target: str               # the target that was checked

    def __str__(self) -> str:
        backends = ", ".join(sorted(self.declared))
        return (
            f"{self.func}: {self.kind}_backends={{{backends}}} "
            f"excludes target {self.target!r} — will fall back"
        )


def _iter_funccalls(node: Any, seen: Set[int] | None = None) -> Iterator[FuncCall]:
    """Yield every :class:`FuncCall` transitively reachable from ``node``.

    Accepts an :class:`Expr`, a Variable (uses ``_expr``), or a
    MultiVariable (walks all child Variables). :class:`VarRef` nodes
    are followed into the referenced Variable's ``_expr`` so composed
    formulas (``NPV(...) + IRR(...)``) surface every underlying call,
    not just the top-level operators.

    ``seen`` tracks Variable identities visited on this walk so cycles
    and diamond deps don't cause infinite recursion or duplicate issues.
    """
    if node is None:
        return
    if seen is None:
        seen = set()
    if isinstance(node, FuncCall):
        yield node
        for arg in node.args:
            if isinstance(arg, Expr):
                yield from _iter_funccalls(arg, seen)
        return
    if isinstance(node, VarRef):
        var = node.var
        if id(var) in seen:
            return
        seen.add(id(var))
        yield from _iter_funccalls(getattr(var, "_expr", None), seen)
        return
    if isinstance(node, Expr):
        # Generic Expr — iter_refs only surfaces Variables, not nested
        # Exprs, so descend into known structural fields instead.
        for child in _expr_children(node):
            yield from _iter_funccalls(child, seen)
        return
    expr = getattr(node, "_expr", None)
    if expr is not None:
        if id(node) in seen:
            return
        seen.add(id(node))
        yield from _iter_funccalls(expr, seen)
    children = getattr(node, "children", None)
    if callable(children):
        for child in children():
            yield from _iter_funccalls(child, seen)


def _expr_children(node: Expr) -> Iterator[Expr]:
    """Yield the direct :class:`Expr` children of a node.

    Each AST node stores its children under a small, fixed set of field
    names. Enumerating them is cheaper than a full reflective walk and
    keeps this module independent of the renderer/walker.
    """
    for attr in ("inner", "left", "right", "operand", "base", "start", "source"):
        child = getattr(node, attr, None)
        if isinstance(child, Expr):
            yield child
    for attr in ("args", "items"):
        for child in getattr(node, attr, ()) or ():
            if isinstance(child, Expr):
                yield child
    for attr in ("kwargs", "variables"):
        for child in (getattr(node, attr, {}) or {}).values():
            if isinstance(child, Expr):
                yield child


def check_compat(
    node: Any,
    target: str,
    *,
    kind: str = "render",
) -> List[CompatIssue]:
    """Return every :class:`FuncCall` whose backend hint excludes ``target``.

    Args:
        node: An :class:`Expr`, Variable, or MultiVariable. The walker
            descends into all reachable FuncCalls.
        target: Backend name to check against (``"excel"``, ``"pandas"``,
            ``"sql"``, etc.). Case-sensitive — match the names used when
            the FuncCall was built.
        kind: ``"render"`` checks :attr:`FuncCall.render_backends`;
            ``"compute"`` checks :attr:`FuncCall.compute_backends`.

    Returns:
        A list of :class:`CompatIssue` — one per offending FuncCall, in
        traversal order. Empty list means the tree is fully compatible.

    A FuncCall whose hint is ``None`` (the default — "universal") is
    always considered compatible and never reported, regardless of
    target.
    """
    if kind not in ("render", "compute"):
        raise ValueError(
            f"kind must be 'render' or 'compute', got {kind!r}."
        )
    attr = f"{kind}_backends"
    issues: List[CompatIssue] = []
    for call in _iter_funccalls(node):
        declared = getattr(call, attr)
        if declared is None:
            continue
        if target not in declared:
            issues.append(CompatIssue(
                func=call.func,
                kind=kind,
                declared=declared,
                target=target,
            ))
    return issues
