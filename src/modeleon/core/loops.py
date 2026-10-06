# SPDX-License-Identifier: Apache-2.0
"""Rows that read each other across periods — a loop broken by ``lag``.

A debt block is the canonical case::

    with m.debt as d:
        d.interest = mo.lag(d.balance) * m.inp.rate
        d.pot      = m.inp.cash - d.interest - m.inp.scheduled
        d.sweep    = mo.IF(d.pot > 0, d.pot, 0.0) * m.inp.share
        d.balance  = mo.lag(d.balance, fill=1000.0) - m.inp.scheduled - d.sweep

Read row by row this is a circle: balance → interest → pot → sweep →
balance. Read period by period it is not — every read of ``balance``
goes through ``lag``, so period *t* needs only period *t - 1*, which is
already known. There is no fixed point to search for, only an order of
evaluation, and that order is what this module supplies. It is the
shape a hand-built workbook has always had: a cell that stands on the
next row's previous column.

Where it applies. Only in a LOOP CONTAINER — ``mo.MultiVariable(...,
loop=True)``, or a class declared ``class X(mo.MultiVariableClass,
loop=True)`` — so nothing changes for any other container::

    m.debt = mo.MultiVariable("Debt", loop=True)
    with m.debt as d:
        ...

How it fits an eager engine. Values compute at operator time, so the
first line above needs ``d.balance`` before it exists. In a loop
container, reading a row it does not have yet returns a PLACEHOLDER — a
series over the model's horizon — so every operator and function builds
its formula tree exactly as usual. The numbers built on a placeholder
are provisional, and every Variable carrying one is marked as WAITING
on it (``Variable.expr`` does the marking); its ``.value`` refuses until
the loop closes. When ``d.balance = …`` finally runs, the placeholder
*becomes* that expression — the same object, so every formula already
points at the real row — and :func:`settle` recomputes the waiting
Variables from their formula trees: the rows that form a cycle period by
period, everything downstream of them as whole series. A loop is built
whole: assigning a row of a loop container a second time is refused —
build the container again (keep its line in the same notebook cell).

A cycle WITHIN one period — a row that needs itself, through other
rows, in the same period — has no order to find; it is refused as
:class:`~modeleon.core.errors.SamePeriodCycleError`, which names its rows.
Such a cycle has no evaluation order: it would need fixed-point
iteration, which the engine does not perform.
"""

from __future__ import annotations

import operator
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Set, Tuple

from .expr import (
    BinOp, Compare, Expr, FuncCall, ListExpr, Literal, MethodCall, Paren,
    SelfRef, Subscript, UnaryOp, VarRef,
)


class Dependents:
    """The Variables built on a placeholder: each once, in the order they
    were built. A list with membership by identity — a row that becomes
    real hands its dependents to every placeholder it waits on, and a
    scan per pair made a block of N rows reading each other cubic in N."""

    __slots__ = ('_items',)

    def __init__(self, items: Any = ()) -> None:
        self._items: Dict[int, Any] = {}
        for item in items:
            self.append(item)

    def append(self, item: Any) -> None:
        self._items.setdefault(id(item), item)

    def __contains__(self, item: Any) -> bool:
        return id(item) in self._items

    def __iter__(self) -> Any:
        return iter(list(self._items.values()))

    def __len__(self) -> int:
        return len(self._items)

    def __reduce__(self) -> Any:
        # Identities do not survive a pickle: rebuild from the items.
        return (Dependents, (list(self._items.values()),))


#: Functions that map one period to one period, so the loop can call
#: them with this period's numbers. Everything else — ``SUM``, ``MAX``,
#: ``NPV``, … — reads a whole series and cannot sit inside the loop.
_ELEMENTWISE = frozenset({
    'IF', 'AND', 'OR', 'NOT', 'ABS', 'ROUND', 'INT', 'MOD', 'CHOOSE',
    'ISBLANK', 'EDATE', 'EOMONTH', 'YEAR', 'MONTH', 'DAY', 'DATE',
    'DAYS360', 'TODAY', 'LEN', 'UPPER', 'LOWER', 'CONCAT',
})

_BINOPS = {
    '+': operator.add, '-': operator.sub, '*': operator.mul,
    '/': operator.truediv, '//': operator.floordiv, '%': operator.mod,
    '**': operator.pow,
}
_COMPARES = {
    '==': operator.eq, '!=': operator.ne, '<': operator.lt,
    '<=': operator.le, '>': operator.gt, '>=': operator.ge,
}


# ─── Naming, for messages ───────────────────────────────────────────


def _label(var: Any) -> str:
    """How a message names a Variable: its row name, else its label,
    else a hint that it is an unnamed intermediate expression."""
    name = getattr(var, '_python_name', None)
    if name:
        owner = getattr(var, '_owner', None)
        owner_name = getattr(owner, '_python_name', None) if owner is not None else None
        return f"{owner_name}.{name}" if owner_name else name
    label = getattr(var, '_display_name', None)
    if label:
        return repr(label)
    expr = getattr(var, '_expr', None)
    if expr is not None:
        text = expr.to_string()
        if len(text) > 40:
            text = text[:37] + "..."
        return f"({text})"
    return "an unnamed row"


def pending_message(var: Any) -> str:
    """Why ``var.value`` is not available yet."""
    error = getattr(var, '_loop_error', None) or next(
        (p._loop_error for p in (getattr(var, '_awaits', None) or ())
         if getattr(p, '_loop_error', None)), None)
    if error:
        return f"{_label(var)} has no value: its loop could not be computed — {error}"
    if any(getattr(p, '_forward_abandoned', False) for p in (getattr(var, '_awaits', None) or ())):
        waits = sorted({_forward_name(p) for p in var._awaits
                        if getattr(p, '_forward_abandoned', False)})
        return (
            f"{_label(var)} has no value: it reads "
            f"{', '.join(repr(w) for w in waits)}, which the block read "
            f"ahead but never assigned."
        )
    forward = getattr(var, '_forward_of', None)
    if forward is not None:
        owner, name = forward
        return (
            f"'{name}' is read before it is assigned: its value is "
            f"available once `{_label(owner)}.{name} = ...` has run."
        )
    waits = sorted(
        {name for p in (getattr(var, '_awaits', None) or ())
         for name in [_forward_name(p)]}
    )
    return (
        f"{_label(var)} is not computed yet: it reads "
        f"{', '.join(repr(w) for w in waits)} before "
        f"{'it is' if len(waits) == 1 else 'they are'} assigned. The value "
        f"is available once the loop closes — the assignment of "
        f"{', '.join(repr(w) for w in waits)} further down."
    )


def _forward_name(placeholder: Any) -> str:
    forward = getattr(placeholder, '_forward_of', None)
    if forward is not None:
        return forward[1]
    return getattr(placeholder, '_python_name', None) or _label(placeholder)


# ─── Reading a formula tree by WHEN it reads ────────────────────────


def _shift_periods(node: MethodCall) -> int:
    periods = node.args[0] if node.args else 1
    try:
        return int(periods)
    except (TypeError, ValueError):
        return 1


def classify_refs(expr: Optional[Expr]) -> Tuple[List[Any], List[Any], List[Any]]:
    """Split a formula's references into ``(same, lagged, led)``: the
    Variables it reads THIS period, LAST period (under a positive
    ``shift`` — ``mo.lag``), and NEXT period (a negative shift — a lead).

    The lagged reads are the ones that break a loop; the same-period
    reads are the ones that order rows within a period.
    """
    same: List[Any] = []
    lagged: List[Any] = []
    led: List[Any] = []

    def walk(node: Any, mode: int) -> None:
        if isinstance(node, VarRef):
            (same if mode == 0 else lagged if mode > 0 else led).append(node.var)
        elif isinstance(node, Paren):
            walk(node.inner, mode)
        elif isinstance(node, (BinOp, Compare)):
            walk(node.left, mode)
            walk(node.right, mode)
        elif isinstance(node, UnaryOp):
            walk(node.operand, mode)
        elif isinstance(node, FuncCall):
            for arg in node.args:
                if isinstance(arg, Expr):
                    walk(arg, mode)
        elif isinstance(node, MethodCall):
            if node.method == 'shift':
                k = _shift_periods(node)
                walk(node.base, 1 if k > 0 else -1 if k < 0 else mode)
            else:
                walk(node.base, mode)
            for arg in node.args:
                if isinstance(arg, Expr):
                    walk(arg, mode)
            for value in node.kwargs.values():
                if isinstance(value, Expr):
                    walk(value, mode)
        elif node is not None:
            for var in node.iter_refs():
                same.append(var)

    walk(expr, 0)
    return same, lagged, led


def same_period_refs(expr: Optional[Expr]) -> List[Any]:
    """The Variables a formula reads in the SAME period — its
    dependency edges for ordering. A read through ``lag`` crosses a
    period boundary and is deliberately not among them."""
    same, _lagged, led = classify_refs(expr)
    return same + led


# ─── Settling a loop once its last row is assigned ─────────────────


def settle(placeholder: Any) -> None:
    """Recompute everything that waited on ``placeholder``, now that it
    has become a real row. A no-op while some Variable in the set still
    waits on ANOTHER row that has not been assigned yet — the whole set
    settles when the last of them lands."""
    dependents = getattr(placeholder, '_forward_dependents', None) or []
    waiting = [v for v in dependents if getattr(v, '_awaits', None)]
    if getattr(placeholder, '_awaits', None):
        waiting.append(placeholder)
    ready: List[Any] = []
    seen: Set[int] = set()
    for v in waiting:
        if id(v) in seen:
            continue
        seen.add(id(v))
        if all(getattr(p, '_forward_of', None) is None for p in v._awaits):
            ready.append(v)
    if not ready:
        return
    involved: Set[Any] = {placeholder}
    for v in ready:
        involved.update(v._awaits)
    _settle_nodes(ready, getattr(placeholder, '_forward_horizon', None))
    # Every placeholder the settled rows waited on lets go of them.
    for p in involved:
        dependents_of_p = getattr(p, '_forward_dependents', None) or []
        if all(getattr(v, '_awaits', None) is None for v in dependents_of_p):
            p._forward_dependents = None
            p._forward_horizon = None


def _refs(v: Any) -> List[Any]:
    return list(getattr(v, '_dependency_refs', None) or ())


def _closure(starts: List[Any]) -> List[Any]:
    """Every Variable reachable from ``starts`` through their references —
    rows and the unnamed intermediates between them."""
    seen: Set[int] = set()
    out: List[Any] = []
    stack = list(starts)
    while stack:
        v = stack.pop()
        if id(v) in seen:
            continue
        seen.add(id(v))
        out.append(v)
        stack.extend(_refs(v))
    return out


def _dependents_of(row: Any, universe: List[Any]) -> List[Any]:
    """The Variables of ``universe`` that read ``row``, directly or through
    others — ``row`` itself included when it reads itself."""
    readers: Dict[int, List[Any]] = {}
    for v in universe:
        for ref in _refs(v):
            readers.setdefault(id(ref), []).append(v)
    seen: Set[int] = set()
    out: List[Any] = []
    stack = list(readers.get(id(row), ()))
    while stack:
        v = stack.pop()
        if id(v) in seen:
            continue
        seen.add(id(v))
        out.append(v)
        stack.extend(readers.get(id(v), ()))
    return out


def loop_rows(var: Any) -> List[Any]:
    """The named rows on the loop through ``var``: everything ``var``
    reads that also reads ``var`` back."""
    universe = _closure([var])
    back = {id(v) for v in _dependents_of(var, universe)}
    return [v for v in universe
            if (id(v) in back or v is var) and getattr(v, '_python_name', None)]


def _settle_nodes(nodes: List[Any], horizon: Optional[int]) -> None:
    """Compute ``nodes`` from their formula trees: the rows on a cycle
    period by period, the rest as whole series, units after values. A
    failure is kept on the rows it leaves without values, so reading one
    says what went wrong instead of promising the loop will close."""
    try:
        _resolve(nodes, horizon)
    except BaseException as exc:
        for v in nodes:
            if getattr(v, '_awaits', None):
                v._loop_error = str(exc)
        raise


def _resolve(nodes: List[Any], horizon: Optional[int]) -> None:
    nodeset: Set[Any] = set(nodes)
    same: Dict[Any, List[Any]] = {}
    lagged: Dict[Any, List[Any]] = {}
    led: Dict[Any, List[Any]] = {}
    for v in nodes:
        s, lg, ld = classify_refs(v._expr)
        same[v] = [r for r in s if r in nodeset]
        lagged[v] = [r for r in lg if r in nodeset]
        led[v] = [r for r in ld if r in nodeset]

    # The rows that genuinely roll forward together: every node on a
    # cycle of the full graph (same-period + lagged edges), plus
    # everything those rows read that is itself still unsettled.
    # Everything else sits downstream of the loop, and recomputes as a
    # whole series once the loop is done.
    cyclic = _cyclic_nodes(nodes, same, lagged)
    periodic: Set[Any] = set()
    stack = list(cyclic)
    while stack:
        v = stack.pop()
        if v in periodic:
            continue
        periodic.add(v)
        stack.extend(same[v])
        stack.extend(lagged[v])
    downstream = [v for v in nodes if v not in periodic]

    # A row of the loop cannot read another unsettled row one period
    # AHEAD: that period is computed from this one. A row downstream of
    # the loop can — by then the whole series exists.
    for v in periodic:
        if led[v]:
            raise ValueError(
                f"{_label(v)} reads {', '.join(_label(r) for r in led[v])} "
                f"one period AHEAD, but that row is still being computed "
                f"period by period with this one — a lead cannot point into "
                f"a loop. A lead is fine into an input."
            )

    if periodic:
        order = _order(periodic, same)
        n = horizon or _horizon_of(order)
        _evaluate_periodic(order, n)
        for v in order:
            v._awaits = None
            v._loop_error = None
        # The loop's own units: its rows were built on placeholders,
        # which carry none, so they are derived again from the formulas —
        # to a fixed point, as the rows depend on each other.
        _rederive_units(order)
    if downstream:
        deps = {v: same[v] + lagged[v] + led[v] for v in downstream}
        # A downstream temporary that no row holds (an expression built and
        # thrown away, or one whose own construction failed) must not stop
        # the loop from closing; it keeps its error for whoever reads it.
        held: Set[Any] = set()
        for v in nodes:
            if getattr(v, '_owner', None) is not None:
                held.add(v)
        stack = list(held)
        while stack:
            v = stack.pop()
            for r in same.get(v, []) + lagged.get(v, []) + led.get(v, []):
                if r not in held:
                    held.add(r)
                    stack.append(r)
        for v in _order(set(downstream), deps):
            try:
                value, unit = _eval_row(v._expr)
            except (ValueError, TypeError) as exc:
                if v in held:
                    raise ValueError(f"{_label(v)}: {exc}") from exc
                v._loop_error = str(exc)
                continue
            _write_back(v, value)
            if not getattr(v, '_unit_declared', False):
                v._unit = unit
            v._awaits = None
            v._loop_error = None


# ─── Units ─────────────────────────────────────────────────────────


def _unit_of(node: Any) -> Any:
    """The unit a formula tree produces, by the rules the operators use
    (``_unit_add`` / ``_unit_mul`` / …) — so a mix that the operators
    refuse is refused here with the same message."""
    from .variable_ops import _VariableArithmetic as arith

    if isinstance(node, Literal):
        return None
    if isinstance(node, VarRef):
        return getattr(node.var, '_unit', None)
    if isinstance(node, Paren):
        return _unit_of(node.inner)
    if isinstance(node, UnaryOp):
        return _unit_of(node.operand)
    if isinstance(node, BinOp):
        left, right = _unit_of(node.left), _unit_of(node.right)
        if node.op in ('+', '-'):
            return arith._unit_add(left, right)
        if node.op == '*':
            return arith._unit_mul(left, right)
        if node.op in ('/', '//'):
            return arith._unit_div(left, right)
        # ``%`` and ``**`` track no unit on the operators either.
        return None
    if isinstance(node, Compare):
        return None
    if isinstance(node, FuncCall):
        from ..functions._helpers import _UNITLESS_RESULTS, _shared_arg_unit
        if node.func in _UNITLESS_RESULTS:
            return None
        return _shared_arg_unit(node.args)
    if isinstance(node, MethodCall) and node.method in ('shift', 'copy'):
        return _unit_of(node.base)
    return None


def _rederive_units(nodes: List[Any]) -> None:
    """Units were derived while the loop's rows were placeholders, which
    carry none. Re-derive them from the final formulas — to a fixed
    point, because a loop's units depend on each other the way its
    values do. A unit the author declared stays as declared: a label on
    a pure number does not survive a product, so it cannot be derived."""
    from .unit import unit_key

    def key(u: Any) -> Any:
        return None if u is None else unit_key(u)

    rows = [v for v in nodes if getattr(v, '_expr', None) is not None]
    declared = {id(v) for v in rows if getattr(v, '_unit_declared', False)}
    for _pass in range(len(rows) + 1):
        changed = False
        for v in rows:
            # Every row's formula is checked - a declared unit decides what
            # the row prints, never what its formula may add.
            try:
                unit = _unit_of(v._expr)
            except ValueError as exc:
                raise ValueError(f"{_label(v)}: {exc}") from exc
            if id(v) in declared:
                continue
            # By key, not equality: a pure number's label is part of what
            # the row prints, though every pure number equals every other.
            if key(unit) != key(v._unit):
                v._unit = unit
                changed = True
        if not changed:
            return


def _cyclic_nodes(nodes: List[Any], same: Dict[Any, List[Any]],
                  lagged: Dict[Any, List[Any]]) -> Set[Any]:
    """Nodes that lie on a cycle of the full dependency graph (Tarjan's
    strongly connected components; a self-edge counts)."""
    index: Dict[Any, int] = {}
    low: Dict[Any, int] = {}
    on_stack: Set[Any] = set()
    stack: List[Any] = []
    cyclic: Set[Any] = set()
    counter = [0]

    def edges(v: Any) -> List[Any]:
        return same[v] + lagged[v]

    def strongconnect(root: Any) -> None:
        # Iterative Tarjan — a long chain of intermediates must not
        # recurse past Python's limit.
        work = [(root, iter(edges(root)))]
        index[root] = low[root] = counter[0]
        counter[0] += 1
        stack.append(root)
        on_stack.add(root)
        while work:
            v, it = work[-1]
            advanced = False
            for w in it:
                if w not in index:
                    index[w] = low[w] = counter[0]
                    counter[0] += 1
                    stack.append(w)
                    on_stack.add(w)
                    work.append((w, iter(edges(w))))
                    advanced = True
                    break
                if w in on_stack:
                    low[v] = min(low[v], index[w])
            if advanced:
                continue
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[v])
            if low[v] == index[v]:
                component: List[Any] = []
                while True:
                    w = stack.pop()
                    on_stack.discard(w)
                    component.append(w)
                    if w is v:
                        break
                # Identity, never ``in``: ``==`` on a Variable is an
                # operator that builds a comparison row.
                if len(component) > 1 or any(w is v for w in edges(v)):
                    cyclic.update(component)

    for v in nodes:
        if v not in index:
            strongconnect(v)
    return cyclic


def _order(nodes: Set[Any], deps: Dict[Any, List[Any]]) -> List[Any]:
    """Topological order over ``deps`` (edges restricted to ``nodes``);
    a cycle that remains is refused by name."""
    pending: Dict[Any, Set[Any]] = {
        v: {d for d in deps[v] if d in nodes and d is not v} for v in nodes
    }
    self_reads = [v for v in nodes if any(d is v for d in deps[v])]
    if self_reads:
        v = self_reads[0]
        from .errors import SamePeriodCycleError
        raise SamePeriodCycleError(
            [v],
            f"{_label(v)} reads itself in the same period, so it cannot be "
            f"computed. Read last period's value instead: mo.lag(...).",
            nodes=[v],
        )
    ordered: List[Any] = []
    while pending:
        ready = [v for v, ds in pending.items() if not ds]
        if not ready:
            cycle = list(_cyclic_nodes(
                list(pending), {v: list(ds) for v, ds in pending.items()},
                {v: [] for v in pending}))
            rows = [v for v in cycle if getattr(v, '_python_name', None)]
            from .errors import SamePeriodCycleError
            if len(rows) == 1:
                raise SamePeriodCycleError(
                    rows,
                    f"{_label(rows[0])} reads itself in the same period, so it "
                    f"cannot be computed. Read last period's value instead: "
                    f"mo.lag(...).",
                    nodes=cycle,
                )
            names = sorted(_label(v) for v in rows) or [_label(v) for v in cycle]
            raise SamePeriodCycleError(
                rows,
                f"these rows read each other WITHIN one period, so none "
                f"of them can be computed first: {' → '.join(names)} → … "
                f"A loop is fine when every round trip passes through "
                f"mo.lag (last period's value) — this one does not.",
                nodes=cycle,
            )
        for v in ready:
            ordered.append(v)
            del pending[v]
        for ds in pending.values():
            ds.difference_update(ready)
    return ordered


def _horizon_of(nodes: List[Any]) -> int:
    for v in nodes:
        if isinstance(v._value, list):
            return len(v._value)
    raise ValueError("a loop needs a horizon: declare default_periods on the model.")


# ─── Period-by-period evaluation ────────────────────────────────────


def _evaluate_periodic(order: List[Any], n: int) -> None:
    for v in order:
        _check_elementwise(v)
    values: Dict[Any, List[Any]] = {v: [None] * n for v in order}
    for t in range(n):
        for v in order:
            values[v][t] = _eval_at(v._expr, t, values, n)
    for v in order:
        _write_back(v, values[v])


def _check_elementwise(v: Any) -> None:
    expr = v._expr

    def walk(node: Any) -> None:
        if isinstance(node, (Literal, VarRef, Subscript)):
            return
        if isinstance(node, Paren):
            walk(node.inner)
        elif isinstance(node, (BinOp, Compare)):
            walk(node.left)
            walk(node.right)
        elif isinstance(node, UnaryOp):
            walk(node.operand)
        elif isinstance(node, FuncCall):
            if node.func not in _ELEMENTWISE:
                from .. import functions as fns
                if not hasattr(fns, node.func):
                    raise ValueError(
                        f"{_label(v)} applies {node.func} to a row of the loop, "
                        f"and {node.func} cannot be run one period at a time. "
                        f"Build that row after the loop closes."
                    )
                hint = ""
                if node.func in ('MAX', 'MIN'):
                    hint = (" For a per-period floor or cap write "
                            "mo.IF(x > 0, x, 0) / mo.IF(x < cap, x, cap).")
                raise ValueError(
                    f"{_label(v)} applies {node.func} to a row of the loop. "
                    f"{node.func} reads a whole series, and inside the loop "
                    f"the series is still being computed.{hint}"
                )
            for arg in node.args:
                if isinstance(arg, Expr):
                    walk(arg)
        elif isinstance(node, MethodCall):
            if node.method not in ('shift', 'copy'):
                raise ValueError(
                    f"{_label(v)}: .{node.method}() is not supported inside "
                    f"a loop yet."
                )
            walk(node.base)
            for value in node.kwargs.values():
                if isinstance(value, Expr):
                    walk(value)
        else:
            raise ValueError(
                f"{_label(v)}: {type(node).__name__} expressions are not "
                f"supported inside a loop yet."
            )

    walk(expr)


def _outside_at(var: Any, t: int, n: int) -> Any:
    """A settled Variable's value at period ``t``."""
    x = var._value
    from .tracks import TrackValues
    if isinstance(x, TrackValues):
        raise ValueError(
            f"{_label(var)} carries tracks; tracks inside a loop are not "
            f"supported yet."
        )
    if isinstance(x, list):
        if t < len(x):
            return x[t]
        if len(x) == 1:
            return x[0]
        raise ValueError(
            f"{_label(var)} has {len(x)} value(s), the loop runs over {n} "
            f"periods."
        )
    return x


def _eval_at(node: Any, t: int, values: Dict[Any, List[Any]], n: int) -> Any:
    from .variable import Variable
    from .variable_ops import _VariableArithmetic as arith

    if isinstance(node, Literal):
        x = node.value
        if isinstance(x, list):
            # A raw list operand is a series, one value per period —
            # exactly as the operators broadcast it outside a loop.
            if t < len(x):
                return x[t]
            if len(x) == 1:
                return x[0]
            raise ValueError(
                f"a list of {len(x)} value(s) inside a loop that runs over "
                f"{n} periods."
            )
        return x
    if isinstance(node, VarRef):
        var = node.var
        if var in values:
            return values[var][t]
        return _outside_at(var, t, n)
    if isinstance(node, Subscript):
        if any(ref in values for ref in node.iter_refs()):
            raise ValueError(
                "a row of the loop indexed at a fixed period, inside the "
                "loop, is not supported yet."
            )
        return _eval_whole(node)
    if isinstance(node, Paren):
        return _eval_at(node.inner, t, values, n)
    if isinstance(node, UnaryOp):
        x = _eval_at(node.operand, t, values, n)
        if node.op == '-':
            return arith._broadcast_operation(0, x, operator.sub, op_symbol='-')
        return x
    if isinstance(node, BinOp):
        left = _eval_at(node.left, t, values, n)
        right = _eval_at(node.right, t, values, n)
        return arith._broadcast_operation(
            left, right, _BINOPS[node.op],
            safe_divide=(node.op == '/'), op_symbol=node.op,
        )
    if isinstance(node, Compare):
        left = _eval_at(node.left, t, values, n)
        right = _eval_at(node.right, t, values, n)
        return arith._broadcast_operation(
            left, right, _COMPARES[node.op], op_symbol=node.op,
        )
    if isinstance(node, FuncCall):
        args = [
            _eval_at(a, t, values, n) if isinstance(a, Expr) else a
            for a in node.args
        ]
        out = _function(node.func)(*args)
        return out._value if isinstance(out, Variable) else out
    if isinstance(node, MethodCall):
        if node.method == 'copy':
            return _eval_at(node.base, t, values, n)
        # shift — checked by _check_elementwise
        k = _shift_periods(node)
        fill = node.kwargs.get('fill_value', 0.0)
        if isinstance(fill, Expr):
            fill = _eval_at(fill, t, values, n)
        base = node.base
        while isinstance(base, Paren):
            base = base.inner
        if not isinstance(base, VarRef):
            raise ValueError("a shift inside a loop must read a row directly.")
        idx = t - k
        var = base.var
        if var in values:
            got = values[var][idx] if 0 <= idx < n else fill
        else:
            x = var._value
            if isinstance(x, list):
                got = x[idx] if 0 <= idx < len(x) else fill
            else:
                got = x
        # A lagged blank reads as 0, as mo.lag reads it - kept blank in a
        # line of dates or words.
        if got is None:
            from .regrain import hole_value
            got = hole_value(var._value)
        return got
    raise ValueError(f"{type(node).__name__} expressions are not supported inside a loop yet.")


# ─── Whole-series evaluation, for the rows downstream of a loop ─────


def _eval_whole(node: Any) -> Any:
    from .variable import Variable
    from .variable_ops import _VariableArithmetic as arith

    if isinstance(node, Literal):
        return node.value
    if isinstance(node, VarRef):
        return node.var._value
    if isinstance(node, Subscript):
        base = _eval_whole(node.base)
        key = node.key
        if isinstance(key, int) and isinstance(base, list):
            return base[key]
        if isinstance(key, slice) and isinstance(base, list):
            return base[key]
        raise ValueError(
            "a keyed subscript over a row of a loop is not supported yet."
        )
    if isinstance(node, Paren):
        return _eval_whole(node.inner)
    if isinstance(node, UnaryOp):
        x = _eval_whole(node.operand)
        if node.op == '-':
            return arith._broadcast_operation(0, x, operator.sub, op_symbol='-')
        return x
    if isinstance(node, BinOp):
        return arith._broadcast_operation(
            _eval_whole(node.left), _eval_whole(node.right), _BINOPS[node.op],
            safe_divide=(node.op == '/'), op_symbol=node.op,
        )
    if isinstance(node, Compare):
        return arith._broadcast_operation(
            _eval_whole(node.left), _eval_whole(node.right),
            _COMPARES[node.op], op_symbol=node.op,
        )
    if isinstance(node, ListExpr):
        items = [_eval_whole(item) for item in node.items]
        return [x[0] if isinstance(x, list) and len(x) == 1 else x for x in items]
    if isinstance(node, FuncCall):
        args = [_as_argument(a) for a in node.args]
        out = _function(node.func)(*args)
        return out._value if isinstance(out, Variable) else out
    if isinstance(node, MethodCall):
        if node.method == 'copy':
            return _eval_whole(node.base)
        if node.method == 'shift':
            from ..functions.lag import lag
            base = node.base
            while isinstance(base, Paren):
                base = base.inner
            fill = node.kwargs.get('fill_value', 0.0)
            if isinstance(fill, VarRef):
                fill = fill.var
            elif isinstance(fill, Expr):
                fill = _eval_whole(fill)
            if isinstance(base, VarRef):
                return lag(base.var, _shift_periods(node), fill=fill)._value
        raise ValueError(
            f".{node.method}() over a row of a loop is not supported yet."
        )
    if isinstance(node, SelfRef):
        from ..functions.recurrence import recurrence
        variables = {name: _as_argument(e) for name, e in node.variables.items()}
        periods: Any = node.periods
        if isinstance(periods, VarRef):
            periods = periods.var
        elif isinstance(periods, Expr):
            periods = _eval_whole(periods)
        return recurrence(
            _as_argument(node.start), node.template, periods=periods, **variables,
        )._value
    raise ValueError(
        f"{type(node).__name__} expressions over a row of a loop are not "
        f"supported yet."
    )


def _function(name: str) -> Any:
    """The ``mo.*`` function behind a formula's function name."""
    from .. import functions as fns
    fn = getattr(fns, name, None)
    if fn is None:
        raise ValueError(
            f"{name}(...) cannot be computed again when the loop closes — "
            f"it is not a function of the engine. Build the row that uses "
            f"it after the loop's last line."
        )
    return fn


def _eval_row(expr: Any) -> Tuple[Any, Any]:
    """A downstream row's values and unit. A function or a recurrence at
    the top is run again as itself, so its unit is the one it always
    gives; arithmetic follows the operators' unit rules."""
    from .variable import Variable
    if isinstance(expr, FuncCall):
        out = _function(expr.func)(*[_as_argument(a) for a in expr.args])
        if isinstance(out, Variable):
            return out._value, out._unit
        return out, None
    if isinstance(expr, SelfRef):
        value = _eval_whole(expr)
        return value, None
    return _eval_whole(expr), _unit_of(expr)


def _as_argument(node: Any) -> Any:
    """A formula argument as the value a ``mo.*`` function accepts: the
    Variable itself for a reference (so the function sees its unit and
    shape), the constant for a literal."""
    if isinstance(node, VarRef):
        return node.var
    if isinstance(node, Literal):
        return node.value
    if isinstance(node, Expr):
        return _eval_whole(node)
    return node


def _write_back(v: Any, value: Any) -> None:
    v._value = value
    if isinstance(value, list):
        v.var_type = 'list'
        sample = value[0] if value else None
    else:
        sample = value
    if isinstance(sample, bool):
        v.value_type = 'bool'
    elif isinstance(sample, (date, datetime)):
        v.value_type = 'datetime'
