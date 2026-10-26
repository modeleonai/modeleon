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
from .unit import UNSETTLED


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


# Which functions the loop may call with this period's numbers is the
# function registry's to say (``modeleon.functions.registry.admits_in_loop``):
# the element-wise ones, and a call over two or more values that is one
# call per period — ``MAX`` / ``MIN`` over two or more operands (a floor, a
# cap), ``SUM(a, b)``, ``MEDIAN(a, b)``, ``PERCENTILE_INC([a, b, c], k)``.
# Everything else — ``NPV``, ``SUM`` or ``MAX`` over one operand, … — reads
# a whole series and cannot sit inside the loop.

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
        order = _order([v for v in nodes if v in periodic], same)
        n = horizon or _horizon_of(order)
        _evaluate_periodic(order, n)
        for v in order:
            v._loop_error = None
        # The loop's own units: its rows were built on placeholders,
        # which carry none, so they are derived again from the formulas —
        # to a fixed point, as the rows depend on each other. The rows
        # wait until their units settle: a loop refused here leaves them
        # without values, its refusal on each (``_settle_nodes``).
        _rederive_units(order)
        # The check + makes of a count added to a share waited while the
        # loop was open; it is made on the settled lines.
        for v in order:
            try:
                _check_counts(v, order)
            except ValueError as exc:
                raise ValueError(f"{_label(v)}: {exc}") from exc
        for v in order:
            v._awaits = None
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
        for v in _order(downstream, deps):
            try:
                declared = v._unit if getattr(v, '_unit_declared', False) else None
                value, unit = _eval_row(v._expr, declared)
                _check_counts(v)
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


#: No row is being derived: every unit read is the row's own.
_SETTLED: frozenset = frozenset()


def _unit_of(node: Any, unsettled: Any = _SETTLED) -> Any:
    """The unit a formula tree produces, by the rules the operators use
    (``_unit_add`` / ``_unit_mul`` / …) — so a mix that the operators
    refuse is refused here with the same message. While a loop settles,
    a row in ``unsettled`` (ids) reads as :data:`UNSETTLED`: ``+`` takes
    the other side's unit, a product or a power of it stays unsettled,
    and a comparison with it waits."""
    from .variable_ops import _VariableArithmetic as arith

    def unit_of(n: Any) -> Any:
        return _unit_of(n, unsettled)

    if isinstance(node, Literal):
        return None
    if isinstance(node, VarRef):
        if id(node.var) in unsettled:
            return UNSETTLED
        return getattr(_settled(node.var), '_unit', None)
    if isinstance(node, Paren):
        return unit_of(node.inner)
    if isinstance(node, UnaryOp):
        return unit_of(node.operand)
    if isinstance(node, BinOp):
        left, right = unit_of(node.left), unit_of(node.right)
        if node.op in ('+', '-', '%'):
            if left is UNSETTLED or right is UNSETTLED:
                known = right if left is UNSETTLED else left
                return UNSETTLED if known is None or known is UNSETTLED else known
            return (arith._unit_mod if node.op == '%' else arith._unit_add)(left, right)
        if left is UNSETTLED:
            return UNSETTLED
        if node.op == '**':
            exponent = node.right.value if isinstance(node.right, Literal) else None
            return arith._unit_pow(left, exponent)
        if right is UNSETTLED:
            return UNSETTLED
        if node.op == '*':
            return arith._unit_mul(left, right)
        if node.op in ('/', '//'):
            return arith._unit_div(left, right)
        return None
    if isinstance(node, Compare):
        left, right = unit_of(node.left), unit_of(node.right)
        if (left is not None and right is not None and UNSETTLED not in (left, right)
                and left != right):
            from .variable_ops import _incompatible_units
            raise ValueError(_incompatible_units(left, right).replace(
                "Cannot add/subtract", "Cannot compare", 1))
        return None
    if isinstance(node, FuncCall):
        # The function's own law, the one its result was built by; an
        # operand read here is often last period's value (a shift), not a
        # line, so its unit comes from its formula, as for ``+``.
        from ..functions._helpers import function_unit
        return function_unit(node.func, node.args, unit_of,
                             own_law=node.render_backends != frozenset())
    if isinstance(node, MethodCall) and node.method == 'shift':
        # A lag keeps its source's unit; its first cells hold the fill, so a
        # fill in a unit is the source's (a pick, as ``mo.lag`` builds it),
        # and a source not derived yet takes the fill's.
        if unit_of(node.base) is None:
            return None
        from ..functions._helpers import same_unit
        fill = node.kwargs.get('fill_value')
        return same_unit('lag', [node.base, *([fill] if isinstance(fill, Expr) else [])],
                         unit_of)
    if isinstance(node, MethodCall) and node.method == 'copy':
        return unit_of(node.base)
    if isinstance(node, Subscript):
        # A period, a key or a slice of a line is in the line's unit, as
        # ``line[0]`` builds it.
        return unit_of(node.base)
    if isinstance(node, SelfRef):
        # A recurrence is in the unit its template gives from its start,
        # by the one rule ``mo.recurrence`` builds it by - read here with
        # the rows not settled yet; a declared unit= is the row's own, and
        # the row's dimension is checked where that unit is known.
        from ..functions.recurrence import template_unit
        return template_unit(node, unit_of(node.start), None,
                             unsettled=unsettled, check_dimension=False)
    return None


def _rederive_units(nodes: List[Any]) -> None:
    """Units were derived while the loop's rows were placeholders, which
    carry none. Derive them again from the final formulas, to a fixed
    point, as the rows depend on each other: every row starts unsettled,
    so none takes a unit from a placeholder — ``+`` gives the other
    side's unit, a product waits. Rows nothing in the loop settles (a
    flag, an index, a discount factor) are pure numbers, as their
    placeholders were: read so, the derivation goes on, and a product
    with one takes its other factor's unit. Then every row's formula is
    checked with the settled units, and must give the unit its row
    holds. A unit the author declared stays as declared: a label on a
    pure number does not survive a product, so it cannot be derived."""
    rows = [v for v in nodes if getattr(v, '_expr', None) is not None]
    declared = {id(v) for v in rows if getattr(v, '_unit_declared', False)}
    unsettled = {id(v) for v in rows if id(v) not in declared}
    _settle(rows, declared, unsettled)
    if unsettled:
        seeded = set(unsettled)
        for v in rows:
            if id(v) in unsettled:
                v._unit = None
        unsettled.clear()
        _settle(rows, declared, unsettled, seeded)
    # Every row's formula is checked - a declared unit decides what the
    # row prints, never what its formula may add.
    units: Dict[int, Any] = {}
    failed: List[Tuple[Any, ValueError]] = []
    for v in rows:
        try:
            units[id(v)] = _unit_of(v._expr)
        except ValueError as exc:
            failed.append((v, exc))
    if failed:
        v, mix = _the_mix(rows, failed)
        raise ValueError(f"{_label(v)}: {mix}") from mix
    for v in rows:
        unit = units[id(v)]
        if id(v) in declared or _dimension(unit) == _dimension(v._unit):
            continue
        raise ValueError(
            f"{_label(v)}: its unit does not settle - read through mo.lag, its "
            f"formula gives {unit or 'no unit'} from a row in "
            f"{v._unit or 'no unit'}. A row multiplied by a dimension every "
            f"period grows a power of it: give the factor no unit, or "
            f"declare the row's unit (.set_unit(...))."
        )


def _the_mix(rows: List[Any], failed: List[Tuple[Any, ValueError]]) -> Tuple[Any, ValueError]:
    """The failing row whose own formula mixes units. A row that fails
    keeps the unit an earlier pass gave it, and a row that reads it -
    through mo.lag it may be checked first - fails on that unit: read
    every failing row, and what takes its unit from one, as not settled,
    and the mix is the row that still fails."""
    unsettled = {id(v) for v, _exc in failed}
    grew = True
    while grew:
        grew = False
        for v in rows:
            if id(v) in unsettled:
                continue
            try:
                waits = _unit_of(v._expr, unsettled) is UNSETTLED
            except ValueError:
                waits = False
            if waits:
                unsettled.add(id(v))
                grew = True
    for v, _exc in failed:
        try:
            _unit_of(v._expr, unsettled - {id(v)})
        except ValueError as exc:
            return v, exc
    return failed[0]


def _check_counts(v: Any, loop: Optional[List[Any]] = None) -> None:
    """The check ``+`` makes of a bps or pp count added to a % fraction
    (:func:`~modeleon.core.variable_ops._refuse_a_count_added_to_a_share`),
    which waits while an operand's loop is open, made on the settled
    lines: over a row's ``+`` / ``-``, and over the lines a function adds,
    compares or picks (``SUM``, ``MAX``, ``IF``, ``CHOOSE``, …). A side the
    row itself feeds - its row read through mo.lag - holds the row's
    numbers from the second period on, so it is judged by its first, the
    one the row has not written into."""
    from .variable_ops import (_numbers, _refuse_a_count_added_to_a_share,
                               _refuse_counts_among_shares)
    expr = v._expr
    where = None
    pairs: List[Tuple[Any, Any]] = []
    if isinstance(expr, BinOp) and expr.op in ('+', '-'):
        operands = [expr.left, expr.right]
    elif isinstance(expr, FuncCall):
        from ..functions._helpers import _SAME_UNIT_FROM
        start = _SAME_UNIT_FROM.get(expr.func)
        if start is None:
            return
        operands, where = list(expr.args[start:]), f"mo.{expr.func}"
    elif isinstance(expr, MethodCall) and expr.method == 'shift':
        # A lag's fill is its series' first cell: a count beside a fraction.
        fill = expr.kwargs.get('fill_value')
        operands, where = [expr.base, *([fill] if isinstance(fill, VarRef) else [])], "mo.lag"
    elif isinstance(expr, SelfRef):
        # The template's own + and -, {prev} read as the start.
        from ..functions.recurrence import template_count_pairs
        start = _settled(expr.start.var) if isinstance(expr.start, VarRef) else None
        pairs = [(_settled(a), _settled(b)) for a, b in template_count_pairs(expr, start)]
        operands = []
    else:
        return
    sides = []
    for node in operands:
        while isinstance(node, Paren):
            node = node.inner
        if isinstance(node, VarRef):
            sides.append(_settled(node.var))
        elif where is None:
            return
    fed: List[Set[int]] = []

    def numbers(side: Any) -> List[Any]:
        value = side._value
        if loop is not None and isinstance(value, list):
            if not fed:
                fed.append({id(r) for r in _dependents_of(v, loop)})
            if id(side) in fed[0]:
                value = value[:1]
        return _numbers(value)

    for a, b in pairs:
        _refuse_a_count_added_to_a_share(a, b, numbers, where=where)
    _refuse_counts_among_shares(sides, numbers, where=where)


def _filled_source(v: Any, by_id: Dict[int, Any], unsettled: Set[int],
                   seeded: Any, declared: Set[int]) -> Optional[Tuple[Any, Any]]:
    """``(row, unit)`` when ``v`` lags a row of this loop that no formula has
    settled yet (or only as a pure number for want of anything else) with
    a fill whose unit is known: the row is in the fill's unit."""
    expr = v._expr
    if not (isinstance(expr, MethodCall) and expr.method == 'shift'
            and isinstance(expr.base, VarRef)):
        return None
    row = by_id.get(id(expr.base.var))
    if row is None or id(row) in declared:
        return None
    if id(row) not in unsettled and not (id(row) in seeded and row._unit is None):
        return None
    fill = expr.kwargs.get('fill_value')
    if not isinstance(fill, VarRef) or id(fill.var) in unsettled:
        return None
    unit = getattr(_settled(fill.var), '_unit', None)
    return None if unit is None else (row, unit)


def _dimension(unit: Any) -> Dict[str, int]:
    """What a unit measures, its label aside: none is a pure number."""
    return {} if unit is None else unit._unit_components


def _settle(rows: List[Any], declared: Set[int], unsettled: Set[int],
            seeded: Set[int] | frozenset = _SETTLED) -> None:
    """Derive the rows' units to a fixed point, reading a row in
    ``unsettled`` as :data:`UNSETTLED` until its formula settles it; a
    mismatch on the way is left for the check on settled units. A row
    read through ``mo.lag`` with a fill in a unit is in that unit — the
    lag's cells are the fill and the row — so the fill settles it, as
    it does a row ``seeded`` as a pure number for want of anything else."""
    from .unit import unit_key

    def key(u: Any) -> Any:
        return None if u is None else unit_key(u)

    by_id = {id(v): v for v in rows}
    for _pass in range(len(rows) + 1):
        changed = False
        for v in rows:
            source = _filled_source(v, by_id, unsettled, seeded, declared)
            if source is not None:
                row, unit = source
                unsettled.discard(id(row))
                if isinstance(seeded, set):
                    seeded.discard(id(row))
                row._unit = unit
                changed = True
            if id(v) in declared:
                continue
            try:
                unit = _unit_of(v._expr, unsettled)
            except ValueError:
                continue
            if unit is UNSETTLED:
                continue
            # By key, not equality: a pure number's label is part of what
            # the row prints, though every pure number equals every other.
            if id(v) in unsettled or key(unit) != key(v._unit):
                unsettled.discard(id(v))
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


def _order(nodes: List[Any], deps: Dict[Any, List[Any]]) -> List[Any]:
    """Topological order over ``deps`` (edges restricted to ``nodes``),
    rows that do not wait on each other kept in the order given — the
    order they were built in, never their addresses; a cycle that
    remains is refused by name."""
    members: Set[Any] = set(nodes)
    pending: Dict[Any, Set[Any]] = {
        v: {d for d in deps[v] if d in members and d is not v} for v in nodes
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
    _read_outside_by_date(order, values, n)
    for t in range(n):
        for v in order:
            values[v][t] = _eval_at(v._expr, t, values, n)
    for v in order:
        _write_back(v, values[v])


def _read_outside_by_date(order: List[Any], values: Dict[Any, List[Any]], n: int) -> None:
    """The lines outside the loop its rows read that run on windows of
    their own, each materialized onto the loop's window by date
    (:func:`_on_loop_window`) — as the operators read them when the rows
    were built — so every period reads the line's value of its date; a
    line read as it is (on the loop's window, a constant) is left to
    :func:`_outside_at`."""
    home = None
    for v in order:
        home = v.__dict__.get('_owner') or v.__dict__.get('_parent')
        if home is not None:
            break
    from .time import resolve_default_window
    window = resolve_default_window(home) if home is not None else None
    if window is None or window.grain is None or window.start is None:
        return
    from .variable import TimeLoc
    from .variable_ops import _on_loop_window
    loop = TimeLoc(window.start, window.grain, window.first_day)
    seen: Set[int] = set()
    for v in order:
        for ref in v._expr.iter_refs():
            if ref in values or id(ref) in seen:
                continue
            seen.add(id(ref))
            x = _settled(ref)._value
            if not isinstance(x, list):
                continue
            aligned = _on_loop_window(ref, x, v, loop, n)
            if aligned is not None:
                values[ref] = aligned


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
            from ..functions.registry import admits_in_loop
            if not admits_in_loop(node.func, _values_per_period(node)):
                from .. import functions as fns
                if not hasattr(fns, node.func):
                    raise ValueError(
                        f"{_label(v)} applies {node.func} to a row of the loop, "
                        f"and {node.func} cannot be run one period at a time. "
                        f"Build that row after the loop closes."
                    )
                raise ValueError(
                    f"{_label(v)} applies {node.func} to a row of the loop. "
                    f"{node.func} reads a whole series, and inside the loop "
                    f"the series is still being computed.{_per_period_hint(node.func)}"
                )
            for i, arg in enumerate(node.args):
                if isinstance(arg, ListExpr) and _reads_range(node.func, i):
                    for item in arg.items:
                        walk(item)
                elif isinstance(arg, Expr):
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


def _reads_range(func: str, i: int) -> bool:
    """Is argument ``i`` of ``func`` read as a range — every argument of a
    reduction, the positions the registry names of a range function?"""
    from ..functions.registry import is_range_taking, range_positions
    positions = range_positions(func)
    return is_range_taking(func) and (positions is None or i in positions)


def _values_per_period(call: FuncCall) -> int:
    """The values a call combines, the count the registry admits it by:
    the arguments it reads as ranges, a list of rows there counting each
    row — or, reading none, its operands."""
    ranged = [a for i, a in enumerate(call.args) if _reads_range(call.func, i)]
    if not ranged:
        return len(call.args)
    return sum(len(a.items) if isinstance(a, ListExpr) else 1 for a in ranged)


def _per_period_hint(func: str) -> str:
    """How ``func`` is one call per period, when the registry says it can
    be."""
    from ..functions.registry import per_period_arity, range_positions, spec
    s = spec(func)
    if s is None or per_period_arity(func) is None:
        return ""
    if s.shape == 'reduction':
        hint = f" Over two or more operands mo.{func} is one call per period, mo.{func}(a, b)"
        if func in ('MAX', 'MIN'):
            hint += "; a per-period floor or cap is mo.MAX(x, 0) / mo.MIN(x, cap)"
        return hint + "."
    rows = "[a, b]"
    shown = f"{rows}, …" if (range_positions(func) or (0,))[0] == 0 else f"…, {rows}"
    return (f" Over a list of two or more rows mo.{func} is one call per period, "
            f"over each period's cells: mo.{func}({shown}).")


def _settled(var: Any) -> Any:
    """A Variable the loop reads as settled. One whose own recompute
    failed (a temporary no row held when its loop closed) keeps its
    error, never its placeholder numbers: its reader is refused with it."""
    error = getattr(var, '_loop_error', None)
    if error:
        raise ValueError(error)
    return var


def _outside_at(var: Any, t: int, n: int) -> Any:
    """A settled Variable's value at period ``t``."""
    x = _settled(var)._value
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
        from ..functions.registry import is_range_taking
        if is_range_taking(node.func):
            args = [_cell_at(a, t, values, n) for a in node.args]
        else:
            args = [_eval_at(a, t, values, n) if isinstance(a, Expr) else a
                    for a in node.args]
        out = _function(node)(*args)
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
            x = _settled(var)._value
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


def _cell_at(node: Any, t: int, values: Dict[Any, List[Any]], n: int) -> Any:
    """What a call that reads ranges reads of an operand at period ``t``: a
    row as its cell, a list of rows as the range of their cells. A number
    or a date reads the same typed into the call or in a cell; a blank,
    text, a boolean or an error is handed over as a line of that one value,
    so it is read as the workbook reads the cell — never as a value typed
    in (a blank as 0, an error as ``#VALUE!``)."""
    from .variable import Variable
    if isinstance(node, VarRef):
        value = _eval_at(node, t, values, n)
        return Variable(value) if value is None or isinstance(value, (str, bool)) else value
    if isinstance(node, ListExpr):
        return tuple(_cell_at(item, t, values, n) for item in node.items)
    return _eval_at(node, t, values, n) if isinstance(node, Expr) else node


# ─── Whole-series evaluation, for the rows downstream of a loop ─────


def _eval_whole(node: Any, declared_unit: Any = None) -> Any:
    from .variable import Variable
    from .variable_ops import _VariableArithmetic as arith

    if isinstance(node, Literal):
        return node.value
    if isinstance(node, VarRef):
        return _settled(node.var)._value
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
        out = _function(node)(*_call_arguments(node))
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
        # Run again as the row was written: a declared unit= rides along,
        # so the template's dimension check reads it as it did at first.
        line = {'unit': declared_unit} if declared_unit is not None else {}
        return recurrence(
            _as_argument(node.start), node.template, periods=periods, **variables, **line,
        )._value
    raise ValueError(
        f"{type(node).__name__} expressions over a row of a loop are not "
        f"supported yet."
    )


def _function(call: FuncCall) -> Any:
    """The ``mo.*`` function behind a formula's call. A call native to no
    renderer (a ``pyformula``) is not the engine's function of its name,
    whatever the name, and cannot be run again."""
    from .. import functions as fns
    name = call.func
    fn = None if call.render_backends == frozenset() else getattr(fns, name, None)
    if fn is None:
        raise ValueError(
            f"{name}(...) cannot be computed again when the loop closes — "
            f"it is not a function of the engine. Build the row that uses "
            f"it after the loop's last line."
        )
    return fn


def _call_arguments(call: FuncCall) -> List[Any]:
    """A formula call's arguments as its ``mo.*`` function takes them
    (:func:`_as_argument`). An operator joined to a cell - the
    ``CONCAT('>=', cap)`` a ``('>=', cap)`` criteria is written as - goes
    back as that pair, so the function reads the criteria as it was given:
    the cell's unit is compared with its range's and an error the cell
    holds is the result, never the text the join gives."""
    from ..functions.criteria import joined_criteria
    out: List[Any] = []
    for arg in call.args:
        pair = joined_criteria(arg)
        out.append(_as_argument(arg) if pair is None
                   else (pair[0], _as_argument(pair[1])))
    return out


def _eval_row(expr: Any, declared_unit: Any = None) -> Tuple[Any, Any]:
    """A downstream row's values and unit. A function at the top is run
    again as itself, so its unit is the one it always gives; a recurrence,
    a subscript and arithmetic follow the rules :func:`_unit_of` reads
    them by."""
    from .variable import Variable
    if isinstance(expr, FuncCall):
        out = _function(expr)(*_call_arguments(expr))
        if isinstance(out, Variable):
            return out._value, out._unit
        return out, None
    return _eval_whole(expr, declared_unit), _unit_of(expr)


def _as_argument(node: Any) -> Any:
    """A formula argument as the value a ``mo.*`` function accepts: the
    Variable itself for a reference (so the function sees its unit and
    shape), the constant for a literal."""
    if isinstance(node, VarRef):
        return _settled(node.var)
    if isinstance(node, Literal):
        return node.value
    if isinstance(node, ListExpr) and all(isinstance(i, (VarRef, Literal)) for i in node.items):
        # A list of lines and values, as the call was given it: the lines
        # themselves (their units, their cells), never their values.
        return [_as_argument(item) for item in node.items]
    if isinstance(node, Expr):
        return _eval_whole(node)
    return node


def _write_back(v: Any, value: Any) -> None:
    v._value = value
    # Typed by the first cell that holds a value: the placeholder build
    # may have typed the row by an error marker it no longer holds.
    cells = value if isinstance(value, list) else [value]
    if isinstance(value, list):
        v.var_type = 'list'
    sample = next((c for c in cells if c is not None
                   and not (isinstance(c, str) and c.startswith('#'))), None)
    if isinstance(sample, bool):
        v.value_type = 'bool'
    elif isinstance(sample, (date, datetime)):
        v.value_type = 'datetime'
    elif isinstance(sample, (int, float)) and v.value_type not in ('int', 'float'):
        v.value_type = 'float'
