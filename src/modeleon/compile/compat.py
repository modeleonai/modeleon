# SPDX-License-Identifier: Apache-2.0
"""Backend compatibility inspection for an AST, Variable, or MultiVariable.

:func:`check_compat` walks an expression tree (or everything reachable
from a Variable / MultiVariable) and reports every :class:`FuncCall`
whose ``render_backends`` or ``compute_backends`` hint excludes a given
target. An empty list means the tree is fully compatible — a renderer
or compute engine for that target can emit every call natively.

Useful as a pre-flight check before compiling a model to a target
("which calls will land in Excel as values rather than formulas?") and
as data for any tooling that wants to flag fallback-bound nodes.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterator, List, Optional, Set

from ..core.expr import Expr, FuncCall, ListExpr, Literal, VarRef


@dataclass(frozen=True)
class CompatIssue:
    """A :class:`FuncCall` that won't render/compute on a target — or,
    in an Excel workbook, a cell that does not hold its whole formula.

    Emitted by :func:`check_compat` for each call whose backend hint
    excludes the requested target (``kind`` ``"render"`` or
    ``"compute"``), and, for a MultiVariable checked against Excel, for
    each cell written as its computed value (``"value"``) and each
    live formula with a computed number written into it
    (``"inlined"``). The caller can render these into a warning, a
    report, or a CI assertion.
    """

    func: str                 # the call; for a cell, the call written as a number,
                              # else the sub-expression itself
    kind: str                 # "render", "compute", "value", "inlined" or "layout"
    declared: frozenset       # the FuncCall's declared backend set (a cell: empty)
    target: str               # the target that was checked
    qpath: Optional[str] = None     # a cell: the line's qualified path
    period: Optional[int] = None    # a cell: its period index
    inlined: Optional[str] = None   # a cell: the sub-expression written as a number
    literal: Optional[str] = None   # an inlined cell: the number written in its place

    def __str__(self) -> str:
        if self.kind == "layout":
            return self.func
        if self.kind == "inlined":
            return (
                f"{self.qpath}[{self.period}]: {self.inlined} is written as "
                f"{self.literal} inside a live formula"
            )
        if self.kind == "value":
            why = f" — no formula stands for {self.inlined}" if self.inlined else ""
            return f"{self.qpath}[{self.period}]: written as its value{why}"
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
        target: Backend name to check against (``"excel"``, or any other
            name used in a FuncCall's backend hint). Case-sensitive —
            match the names used when the FuncCall was built.
        kind: ``"render"`` checks :attr:`FuncCall.render_backends`;
            ``"compute"`` checks :attr:`FuncCall.compute_backends`.

    Returns:
        A list of :class:`CompatIssue` — one per offending FuncCall, in
        traversal order. Empty list means the tree is fully compatible.

    A FuncCall whose hint is ``None`` (the default — "universal") is
    always considered compatible and never reported, regardless of
    target.

    A MultiVariable checked against ``"excel"`` is also laid out and
    rendered as ``to_excel`` would, and every cell that does not hold
    its whole formula follows: a cell written as its computed value
    (``kind="value"``), and a live formula with a computed number
    written into it (``kind="inlined"`` — ``SUM(price * volume)`` over
    a line the book has no cells for), each with the line, the period
    and the sub-expression. Inputs and authored values are not issues.
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
    if target == "excel" and kind == "render":
        issues.extend(_excel_cell_issues(node, target))
    return issues


def _excel_cell_issues(root: Any, target: str) -> List[CompatIssue]:
    """The cells of ``root``'s workbook that do not hold their whole
    formula — laid out and rendered exactly as ``to_excel`` does, with
    nothing written. A bare Variable or expression has no workbook."""
    from ..core.multi_variable import MultiVariableBase
    from ..core.variable import Variable

    if not isinstance(root, MultiVariableBase):
        return []
    from .excel.layout import LayoutEngine
    from .excel.translator import ExcelTranslator
    from .excel.view import resolve_excel_view
    from .excel.writer import (
        _build_var_to_sheet,
        _collect_root_sheets,
        _identity_homes,
        _resolve_var_id,
        _unique_sheet_names,
        book_time_for,
        cell_formula,
        expand_tracked_tree,
    )

    root = expand_tracked_tree(root, mode=resolve_excel_view(root).tracks or "rows")
    roots = _collect_root_sheets(root)
    if not roots:
        return []
    engine = LayoutEngine(roots, flatten_nested_sheets=True, header_rows=True)
    addresses = engine.compute_addresses()
    if not addresses:
        return []
    sheet_names = _unique_sheet_names(roots)
    var_to_sheet = _build_var_to_sheet(addresses, roots, sheet_names)
    translator = ExcelTranslator(
        addresses, var_to_sheet,
        identity_cells=_identity_homes(roots, addresses),
        time=book_time_for(engine, roots, sheet_names, addresses, var_to_sheet),
    )
    issues: List[CompatIssue] = []
    issues.extend(
        CompatIssue(func=msg, kind="layout", declared=frozenset(), target=target)
        for msg in engine.refusals
    )
    for sheet_mv, sheet_name in zip(roots, sheet_names):
        for _name, comp in sheet_mv.walk():
            if not isinstance(comp, Variable):
                continue
            var_id = _resolve_var_id(comp, addresses)
            if var_id is None:
                continue
            addr = addresses[var_id]
            for i in range(len(addr.values)):
                if not _tries_a_formula(comp, i):
                    continue
                formula = cell_formula(comp, i, var_id, translator, sheet_name, addr)
                if formula is None:
                    refused = getattr(translator, "last_render_lossy_terms", [])
                    issues.append(_cell_issue(
                        "value", target, var_id, i, refused[0] if refused else None, None))
                    continue
                for node, literal in getattr(translator, "last_render_inlined_terms", []):
                    issues.append(_cell_issue("inlined", target, var_id, i, node, literal))
    return issues


def _tries_a_formula(var: Any, period_idx: int) -> bool:
    """Whether the writer writes a formula here before any fallback —
    ``cell_formula``'s own rule: a line with an expression, except a
    positional list whose item at this period is an authored literal."""
    expr = getattr(var, "_expr", None)
    if expr is None:
        shown = getattr(var, "shown_track", None)
        found = shown() if callable(shown) else None
        if found is None:
            return False
        expr = found[1]
    if isinstance(expr, ListExpr):
        items = expr.items
        item = items[period_idx] if 0 <= period_idx < len(items) else (
            items[-1] if items else None
        )
        return item is not None and not isinstance(item, Literal)
    return True


def _cell_issue(
    kind: str, target: str, qpath: str, period: int, node: Any, literal: Optional[str],
) -> CompatIssue:
    text = _spelled(node)
    func = node.func.upper() if isinstance(node, FuncCall) else text
    return CompatIssue(
        func=func or "", kind=kind, declared=frozenset(), target=target,
        qpath=qpath, period=period, inlined=text, literal=literal,
    )


def _spelled(node: Any) -> Optional[str]:
    """The sub-expression as the model spells it; one too deep to spell
    is named by its kind."""
    if node is None:
        return None
    try:
        return str(node.to_string())
    except RecursionError:
        return type(node).__name__
