# SPDX-License-Identifier: Apache-2.0
"""Projection — re-grain a whole model onto a coarser grain.

``model.at(grain)`` walks the tree and re-grains every line:

- a **time-agnostic** constant broadcasts unchanged;
- a line **with a rule** applies it to its own value (``Variable.at`` —
  apply-own-rule); ``mo.ratio(num, den)`` resolves its sibling sources here
  and re-grains as sum-over-sum;
- a **formula without a rule** recomputes over its re-grained inputs — the
  formula's expression is re-run through the *same operators* (the operator
  chokepoint, not a parallel evaluator), which is correct for additive forms
  (``gross = revenue - cogs``)
  and for quotients (``margin = gp / rev`` recomputes as ``Σgp / Σrev`` —
  ratio-of-sums, the financially correct blend);
- a **product of two time-varying lines** (``revenue = price * seats``) can't
  recompute from aggregates (``Σ(p*s) != aggP * aggS``) — its *output* is a
  flow, so the projection evaluates it at native grain and re-grains the
  result by ``sum`` (evaluate-then-regrain). No rule, no error, correct
  numbers for a flow such as price × quantity; declare a rule only to
  override the flow default;
- a **formula without a rule whose total turns on the period's number or
  dates** (``mo.time.index`` / ``.start`` / ``.end``) — arithmetic on them,
  a product of schedules of them, a number over one, a condition or a power
  on them when the condition yields fixed numbers, a level or a schedule —
  has no default: it needs a rule, and a line computed from it without one
  names it.

A shared cache re-grains each source once, so cross-line references agree.
"""

from __future__ import annotations

import math
import operator
from typing import Any, Dict, List, Optional, Tuple

from .expr import (
    BinOp, Compare, Expr, FuncCall, ListExpr, Literal, MethodCall, Paren, Regrain,
    Restrict, SelfRef, Subscript, UnaryOp, VarRef,
)
from .variable import Variable

_BINOPS: Dict[str, Any] = {'+': operator.add, '-': operator.sub,
                           '*': operator.mul, '/': operator.truediv,
                           '**': operator.pow}
_UNARYOPS: Dict[str, Any] = {'-': operator.neg, '+': operator.pos}
_COMPARES: Dict[str, Any] = {'==': operator.eq, '!=': operator.ne,
                             '<': operator.lt, '<=': operator.le,
                             '>': operator.gt, '>=': operator.ge}
#: mo.* functions safe to re-run over re-grained inputs — Variable-aware,
#: element-wise or scalar-broadcasting, no self-reference. Reductions over
#: one series (SUM/NPV/IRR/…) never reach recompute: scalar-valued formulas
#: copy.
_RECOMPUTE_FUNCS = ('IF', 'AND', 'OR', 'NOT', 'CHOOSE',
                    'MAX', 'MIN', 'ABS', 'ROUND', 'INT', 'MOD')
#: Aggregates whose call over two or more operands with a series among
#: them is one per period, re-grained as their operators are: ``SUM`` and
#: ``AVERAGE`` as ``a + b`` — the call over the re-grained operands, each
#: by its own rule (a balance its last month, a flow its sum) — and
#: ``PRODUCT`` as ``a * b``, a flow evaluated at native grain when two of
#: its operands vary in time.
_PER_PERIOD_AGGREGATES = ('SUM', 'AVERAGE', 'PRODUCT')


#: The time fields that say WHICH period a formula is computed for — its
#: number and its dates — as opposed to how long the period is (``days``,
#: ``months``). Recomputed at a coarser grain they are the quarter's or the
#: year's own number and dates, and summed they turn a level into a flow: a
#: formula whose total turns on them has no total without a re-grain rule.
_POSITION_FIELDS = frozenset({'index', 'start', 'end'})

#: The projection cache's slot for what the position law found, by row.
_POSITION_MEMO = object()

#: The projection cache's slot for the lines refused for the period's
#: number or dates, ``{id(line): (message, line, source)}``: met again, a
#: refusal is raised again at once, not worked out anew down the line it
#: came from (and the cache keeps no traceback alive).
_REFUSED = object()

#: What a line refused for the period's number or dates reads, and the rules
#: it can take.
_DATES = "the period's number or dates (mo.time.index, .start or .end)"
_RULES = (".set_regrain(mo.up('sum')) for a flow or a 0/1 flag, "
          ".set_regrain(mo.up('mean')) for a level such as a price, "
          "mo.up('last') for a balance, mo.ratio('num', 'den') for a ratio")

#: Marks a running total built inline on the period's number, dates or
#: length (``mo.cumsum(mo.time.days)``), which counts the periods.
_ROLLED = object()

#: How a value adds up over the periods of a coarser one: a fixed number, a
#: flow (its total is its sum), a level (it is not), a schedule of the
#: period's number or dates (either, as the author says), or not known.
_CONST, _FLOW, _LEVEL, _SCHEDULE, _ANY = 'const', 'flow', 'level', 'schedule', 'any'

#: Where a node of a formula sits for :func:`_position_cause`: in arithmetic
#: the projection re-runs over re-grained inputs, or under a power.
_RECOMPUTED, _POWER = 'recomputed', 'power'


def position_cache(memo: Dict[Any, Any]) -> Dict[int, Any]:
    """A fresh projection cache that keeps what the position law found in
    ``memo`` — for a caller that projects line by line with a cache of its
    own each time (the totals of a tracked line), so the rows a line reads
    are judged once, not once per line."""
    return {id(_POSITION_MEMO): memo}


def _line_label(var: Any) -> str:
    return var._display_name or var.python_name or var.id


def _labels(line: Any, other: Any) -> Tuple[str, str]:
    """The labels of two lines, by their paths when they read alike (an
    alias keeps its source's name)."""
    a, b = _line_label(line), _line_label(other)
    return (line.id, other.id) if a == b and line is not other else (a, b)


def _position_refusal(message: str, line: Any, source: Any) -> ValueError:
    """The refusal of a line with no re-grain rule whose total turns on the
    period's number or dates — a ValueError that carries ``line`` (refused)
    and ``source`` (the line whose own formula makes it so)."""
    error = ValueError(message)
    error.__dict__.update(position_line=line, position_source=source)
    return error


def _is_position_refusal(error: BaseException) -> bool:
    return error.__dict__.get('position_line') is not None


def _time_row_field(var: Any) -> Optional[str]:
    """The ``mo.time`` field ``var`` IS — a time row, a copy or alias of one,
    or one shifted by ``mo.lag`` (the previous period's end is still the
    time of a period) — or None."""
    from .time import time_field_of
    seen: set = set()
    while var is not None and id(var) not in seen:
        seen.add(id(var))
        field = time_field_of(var)
        if field is not None:
            return field
        expr = getattr(var, '_expr', None)
        if (isinstance(expr, MethodCall) and expr.method == 'shift'
                and isinstance(expr.base, VarRef)):
            var = expr.base.var
            continue
        return None
    return None


def _date_of(node: Any) -> Optional[str]:
    """``'start'`` / ``'end'`` when ``node`` reads one of the period's dates
    as it is (a time row of it, an alias, a lag), else None."""
    from .expr import TimeRef
    while isinstance(node, Paren):
        node = node.inner
    if isinstance(node, TimeRef):
        return node.field if node.field in ('start', 'end') else None
    if isinstance(node, VarRef):
        field = _time_row_field(node.var)
        return field if field in ('start', 'end') else None
    return None


def _is_series(value: Any) -> bool:
    from .tracks import TrackValues
    return isinstance(value, list) or (
        isinstance(value, TrackValues) and value.time_length is not None)


def _inline_series(var: Any) -> Any:
    """The formula of ``var`` when it is a series built inline in another
    line's formula (no owner of its own), else None."""
    sub = getattr(var, '_expr', None)
    if (sub is not None and getattr(var, '_owner', None) is None
            and _is_series(getattr(var, '_value', None))):
        return sub
    return None


def _signed_terms(expr: Any) -> List[Tuple[int, Any]]:
    """``expr`` as an additive chain of signed terms — through ``+``, ``-``,
    a unary minus, parentheses and the series built inline — walked with a
    stack, so a chain of any length (a Python ``sum()`` over many lines)
    flattens."""
    out: List[Tuple[int, Any]] = []
    seen: set = set()
    stack: List[Tuple[int, Any]] = [(1, expr)]
    while stack:
        sign, node = stack.pop()
        while isinstance(node, Paren):
            node = node.inner
        if isinstance(node, BinOp) and node.op in ('+', '-'):
            stack.append((sign if node.op == '+' else -sign, node.right))
            stack.append((sign, node.left))
            continue
        if isinstance(node, UnaryOp) and node.op == '-':
            stack.append((-sign, node.operand))
            continue
        if isinstance(node, VarRef) and _date_of(node) is None:
            sub = _inline_series(node.var)
            if isinstance(sub, (BinOp, UnaryOp, Paren)) and id(node.var) not in seen:
                seen.add(id(node.var))
                stack.append((sign, sub))
                continue
        out.append((sign, node))
    return out


def _sub_exprs(node: Any) -> List[Any]:
    """The expressions directly under ``node`` (not through Variables)."""
    if isinstance(node, (BinOp, Compare)):
        return [node.left, node.right]
    if isinstance(node, UnaryOp):
        return [node.operand]
    if isinstance(node, Paren):
        return [node.inner]
    if isinstance(node, FuncCall):
        return [a for a in node.args if isinstance(a, Expr)]
    if isinstance(node, MethodCall):
        return [node.base, *(a for a in node.args if isinstance(a, Expr)),
                *(v for v in node.kwargs.values() if isinstance(v, Expr))]
    if isinstance(node, (Subscript, Restrict)):
        return [node.base]
    if isinstance(node, ListExpr):
        return [item for item in node.items if isinstance(item, Expr)]
    return []


def _parts(node: Any) -> List[Any]:
    """What a walk of a formula looks into under ``node``: its expressions,
    or — for a node that holds rows instead (a projection ``.at(...)``, a
    recurrence or a running total) — a reference to each row it reads."""
    children = _sub_exprs(node)
    if children or not hasattr(node, 'iter_refs'):
        return children
    return [VarRef(var) for var in node.iter_refs()]


def _dates_cancel(terms: List[Tuple[int, Any]]) -> bool:
    """Do an additive chain's reads of the period's dates cancel (``end -
    start``, ``1 + end - start``) — how long the period is, measured on its
    dates, recomputed at a coarser grain as that period's length?"""
    dates = [sign for sign, term in terms if _date_of(term) is not None]
    return len(dates) >= 2 and sum(dates) == 0


def _beyond_duration(terms: List[Tuple[int, Any]]) -> List[Any]:
    """The terms of an additive chain a walk looks into: all of them, or —
    when its dates cancel — those besides the dates (a schedule added to the
    period's length still turns on the dates)."""
    if _dates_cancel(terms):
        return [term for _sign, term in terms if _date_of(term) is None]
    return [term for _sign, term in terms]


def _varies(expr: Any) -> bool:
    """:func:`_is_time_varying` with a stack instead of recursion — what
    :func:`_recompute` asks of a product's operands."""
    seen: set = set()
    stack = [expr]
    while stack:
        node = stack.pop()
        if isinstance(node, Literal):
            continue
        if isinstance(node, VarRef):
            var = node.var
            if var.time is not None:
                return True
            sub = getattr(var, '_expr', None)
            if sub is not None and isinstance(var._value, list) and id(var) not in seen:
                seen.add(id(var))
                stack.append(sub)
            continue
        if isinstance(node, Paren):
            stack.append(node.inner)
        elif isinstance(node, UnaryOp):
            stack.append(node.operand)
        elif isinstance(node, BinOp):
            stack.extend((node.left, node.right))
        else:
            return True     # unknown node: assume time-varying, as _recompute does
    return False


def _is_flow_product(node: Any) -> bool:
    """A product of two lines that vary by period — computed period by
    period and summed by :func:`_recompute`, whatever its factors read."""
    if isinstance(node, BinOp) and node.op == '*':
        return _varies(node.left) and _varies(node.right)
    return (isinstance(node, FuncCall) and node.func == 'PRODUCT'
            and _per_period_call(node) and sum(map(_varies, node.args)) >= 2)


def _row_reads(var: Any) -> Tuple[bool, bool, List[Any]]:
    """What ``var`` reads itself: the period's number or dates, stored data
    (a typed series, a track's entered values), and the other rows it is
    computed from. A duration on the dates and the period's length are
    neither; a running total or a recurrence of the period's length counts
    the periods."""
    from .expr import TimeRef
    field = _time_row_field(var)
    if field is not None:
        return field in _POSITION_FIELDS, False, []
    data = getattr(var, '_role_kwargs', None) is not None
    expr = getattr(var, '_expr', None)
    if expr is None:
        return False, data or _is_series(getattr(var, '_value', None)), []
    if isinstance(expr, ListExpr):
        data = True                 # a typed series, some cells linked to inputs
    rolls = isinstance(expr, SelfRef)
    position = False
    rows: List[Any] = []
    seen: set = set()
    stack = [expr]
    while stack:
        node = stack.pop()
        while isinstance(node, Paren):
            node = node.inner
        if isinstance(node, TimeRef):
            position = position or node.field in _POSITION_FIELDS or rolls
            continue
        if isinstance(node, VarRef):
            ref = node.var
            ref_field = _time_row_field(ref)
            if ref_field is not None:
                position = position or ref_field in _POSITION_FIELDS or rolls
            elif id(ref) not in seen and _is_series(getattr(ref, '_value', None)):
                seen.add(id(ref))
                sub = _inline_series(ref)
                if rolls and sub is not None:
                    stack.append(sub)       # what a running total rolls (days / 365)
                else:
                    rows.append(ref)
            continue
        if not isinstance(node, Expr) or isinstance(node, Subscript):
            continue
        if isinstance(node, BinOp) and node.op in ('+', '-'):
            stack.extend(_beyond_duration(_signed_terms(node)))
            continue
        stack.extend(_parts(node))
    return position, data, rows


def _schedule_flags(var: Any, memo: Dict[Any, Any]) -> Tuple[bool, bool]:
    """(reads the period's number or dates, reads stored data) — anywhere in
    what ``var`` is computed from, through any number of lines. Rows that
    read each other (a loop's, a recurrence) share one answer, worked out
    once over the rows they reach, so it does not depend on which line is
    met first."""
    done: Dict[int, Tuple[bool, bool]] = memo.setdefault('flags', {})
    if id(var) in done:
        return done[id(var)]
    order: Dict[int, int] = {}
    low: Dict[int, int] = {}
    found: Dict[int, List[bool]] = {}
    open_rows: set = set()
    component: List[Any] = []
    frames: List[Tuple[Any, List[Any], int]] = []

    looped: set = memo.setdefault('looped', set())
    reads: Dict[int, List[Any]] = {}

    def enter(row: Any) -> None:
        position, data, rows = _row_reads(row)
        order[id(row)] = low[id(row)] = len(order)
        found[id(row)] = [position, data]
        reads[id(row)] = rows
        component.append(row)
        open_rows.add(id(row))
        frames.append((row, rows, 0))

    enter(var)
    while frames:
        row, rows, i = frames[-1]
        if i < len(rows):
            frames[-1] = (row, rows, i + 1)
            nxt = rows[i]
            if id(nxt) in done:
                for k, flag in enumerate(done[id(nxt)]):
                    found[id(row)][k] = found[id(row)][k] or flag
            elif id(nxt) not in order:
                enter(nxt)
            elif id(nxt) in open_rows:
                low[id(row)] = min(low[id(row)], order[id(nxt)])
            continue
        frames.pop()
        parent = frames[-1][0] if frames else None
        if low[id(row)] == order[id(row)]:
            members = []
            while True:
                member = component.pop()
                open_rows.discard(id(member))
                members.append(member)
                if member is row:
                    break
            flags = (any(found[id(m)][0] for m in members),
                     any(found[id(m)][1] for m in members))
            if len(members) > 1 or any(r is row for r in reads[id(row)]):
                looped.update(id(m) for m in members)
            for member in members:
                done[id(member)] = flags
            if parent is not None:
                for k, flag in enumerate(flags):
                    found[id(parent)][k] = found[id(parent)][k] or flag
        elif parent is not None:
            low[id(parent)] = min(low[id(parent)], low[id(row)])
    return done[id(var)]


def _is_schedule(var: Any, memo: Dict[Any, Any]) -> bool:
    """Is ``var`` a schedule of the period's number or dates — a line whose
    values follow from them and from fixed numbers alone (a flag of the
    dates, a counter of the periods, an amount switched on by a date), at
    any distance and whatever its rule? A line that reads stored data
    (``units * price``) is none, though it varies with them."""
    if not _is_series(getattr(var, '_value', None)):
        return False                # one number (a sum over the window) holds
    position, data = _schedule_flags(var, memo)
    return position and not data


def _turns_on_dates(var: Any, memo: Dict[Any, Any]) -> bool:
    """Does a row's value turn on the period's number or dates anywhere in
    what it is computed from — a schedule, or a line that mixes one with
    stored data (``installed * cod_flag``)?"""
    if not _is_series(getattr(var, '_value', None)):
        return False
    return _schedule_flags(var, memo)[0]


def _reads_data(node: Any, memo: Dict[Any, Any]) -> bool:
    """Does ``node`` read stored data — a typed series or entered actuals —
    anywhere in what it is computed from?"""
    seen: set = set()
    stack = [node]
    while stack:
        item = stack.pop()
        if isinstance(item, VarRef):
            var = item.var
            if id(var) in seen or _time_row_field(var) is not None:
                continue
            seen.add(id(var))
            if _is_series(getattr(var, '_value', None)) and _schedule_flags(var, memo)[1]:
                return True
            continue
        if isinstance(item, Expr):
            stack.extend(_parts(item))
    return False


def _var_kind(var: Any, memo: Dict[Any, Any]) -> str:
    """How a row adds up when its formula is not looked into: by its rule
    ('sum' a flow, any other a level), a time row as time, a cell as a fixed
    number, a stored series with no rule as not known."""
    field = _time_row_field(var)
    if field is not None:
        return _SCHEDULE if field in _POSITION_FIELDS else _FLOW
    spec = getattr(var, '_regrain', None)
    if spec is not None:
        rules = [r for r in (spec.default, *spec.overrides.values()) if r is not None]
        if rules and all(r == 'sum' for r in rules):
            return _FLOW
        return _ANY if any(r == 'sum' for r in rules) else _LEVEL
    if not _is_series(getattr(var, '_value', None)):
        return _CONST
    return _ANY


def _zero_one(var: Any) -> bool:
    """Does ``var`` hold ones and zeros only (a flag of the dates, a test)?"""
    from .tracks import TrackValues
    value = getattr(var, '_value', None)
    tracks = list(value.values()) if isinstance(value, TrackValues) else [value]
    if not tracks or not all(isinstance(t, list) for t in tracks):
        return False
    items = [x for t in tracks if isinstance(t, list) for x in t if x is not None]
    return bool(items) and all(isinstance(x, (bool, int, float)) and x in (0, 1)
                               for x in items)


def _is_gate(node: Any) -> bool:
    """Is ``node`` a 0/1 flag — a series of ones and zeros (a flag of the
    dates, a test, a product of flags), or one scaled by a number (``55 *
    flag``) — that turns another factor on and off without changing its
    kind? The period's length is no flag (months is all ones)."""
    while True:
        while isinstance(node, Paren):
            node = node.inner
        if isinstance(node, VarRef):
            if _time_row_field(node.var) is not None:
                return False
            if _zero_one(node.var):
                return True
            sub = _inline_series(node.var)
            if sub is None:
                return False
            node = sub
            continue
        if isinstance(node, BinOp) and node.op in ('*', '/'):
            if not _varies(node.right):
                node = node.left
                continue
            if node.op == '*' and not _varies(node.left):
                node = node.right
                continue
        return False


def _product_kind(node: Any, factors: List[Any], kinds: List[str],
                  memo: Dict[Any, Any]) -> str:
    """The kind of a product of lines that vary by period: one that reads
    stored data is a flow (``units * price``, ``headcount * salary``: the
    flow default of a product); a 0/1 flag is a gate that keeps the kind of
    what it multiplies; a flow or the period's length among the rest makes a
    flow (``gen * tariff``, ``balance * rate * days / 365``), a schedule
    among them a schedule (``55 * flag * indexation``); else not known."""
    if _reads_data(node, memo):
        return _FLOW
    rest = [kind for factor, kind in zip(factors, kinds) if not _is_gate(factor)]
    if _FLOW in rest:
        return _FLOW
    if not rest or _SCHEDULE in rest or _turns_on_position(node, memo) is not None:
        return _SCHEDULE            # a flag times a level of the dates, a tariff
    return _ANY


def _joined(kinds: List[str]) -> str:
    """The kind of values put side by side (the branches of a condition, the
    terms of a sum): a fixed number takes the others' kind."""
    rest = [kind for kind in kinds if kind != _CONST]
    if not rest:
        return _CONST
    if _SCHEDULE in rest:
        return _SCHEDULE
    return rest[0] if all(kind == rest[0] for kind in rest) else _ANY


def _kind_parts(node: Any, memo: Dict[Any, Any]) -> List[Any]:
    """The values a node's kind is made of (a condition's branches, not its
    test)."""
    if isinstance(node, Paren):
        return [node.inner]
    if isinstance(node, VarRef):
        var = node.var
        sub = getattr(var, '_expr', None)
        if (sub is None or _time_row_field(var) is not None
                or getattr(var, '_regrain', None) is not None
                or not _is_series(getattr(var, '_value', None))):
            return []
        if getattr(var, '_owner', None) is not None:
            _schedule_flags(var, memo)          # learns the loops the row is in
            if id(var) in memo.get('looped', ()):
                return []                       # a loop's row: no kind of its own
        return [sub]            # a series built inline, or a formula with no rule
    if isinstance(node, BinOp):
        if node.op in ('+', '-'):
            return _beyond_duration(_signed_terms(node))
        return [node.left, node.right] if node.op in ('*', '/') else []
    if isinstance(node, UnaryOp):
        return [node.operand]
    if isinstance(node, FuncCall):
        args = [a for a in node.args if isinstance(a, Expr)]
        if node.func in ('IF', 'CHOOSE'):
            return [a for a in node.args[1:] if isinstance(a, Expr)]
        if node.func in ('MIN', 'MAX', 'SUM', 'AVERAGE', 'PRODUCT'):
            return args
        if node.func in ('ROUND', 'INT', 'ABS', 'MOD'):
            return args[:1]
        return []
    if isinstance(node, MethodCall) and node.method in ('shift', 'copy'):
        return [node.base]
    if isinstance(node, Restrict):
        return [node.base]          # a track's slice keeps its line's kind
    return []


#: Functions that count the days between dates — the period's length when
#: read on its dates.
_DAY_COUNTS = ('DAYS', 'DAYS360', 'YEARFRAC', 'NETWORKDAYS', 'NETWORKDAYS_INTL')


def _changes(terms: List[Tuple[int, Any]]) -> set:
    """The positions of the terms that pair ``x - mo.lag(x)`` in an additive
    chain: the change in a balance, a flow."""
    def lagged(term: Any) -> Any:
        node = term.var._expr if isinstance(term, VarRef) else term
        if isinstance(node, MethodCall) and node.method == 'shift' and isinstance(node.base, VarRef):
            return node.base.var
        return None
    out: set = set()
    for i, (sign, term) in enumerate(terms):
        base = lagged(term)
        if base is None:
            continue
        for j, (other_sign, other) in enumerate(terms):
            if (j not in out and j != i and other_sign == -sign
                    and isinstance(other, VarRef) and other.var is base):
                out.update((i, j))
                break
    return out


def _node_kind(node: Any, kinds: List[str], memo: Dict[Any, Any]) -> str:
    from .expr import TimeRef
    if isinstance(node, (Literal, Subscript)):
        return _CONST
    if isinstance(node, TimeRef):
        return _SCHEDULE if node.field in _POSITION_FIELDS else _FLOW
    if isinstance(node, VarRef):
        var = node.var
        if not kinds:
            return _var_kind(var, memo)
        if getattr(var, '_owner', None) is None or kinds[0] in (_FLOW, _LEVEL):
            return kinds[0]     # a series built inline, or a formula that adds up as one
        return _SCHEDULE if _is_schedule(var, memo) else _ANY
    if isinstance(node, (Paren, UnaryOp)):
        return kinds[0] if kinds else _ANY
    if isinstance(node, BinOp):
        if node.op in ('+', '-'):
            # a chain whose dates cancel holds the period's length, a flow;
            # x - mo.lag(x) is the change in a balance, a flow
            terms = _signed_terms(node)
            kept = [t for t in terms if not (_dates_cancel(terms) and _date_of(t[1]) is not None)]
            pairs = _changes(kept)
            kinds = [_FLOW if i in pairs else k for i, k in enumerate(kinds)]
            return _joined(kinds + [_FLOW] if _dates_cancel(terms) else kinds)
        if node.op == '*':
            if _is_flow_product(node):
                return _product_kind(node, [node.left, node.right], kinds, memo)
            left, right = kinds
            if _is_gate(node):
                return _CONST       # a number turned on and off by a flag
            return right if left == _CONST else left if right == _CONST else _ANY
        if node.op == '/':
            left, right = kinds
            if right == _CONST:
                return left
            return _LEVEL if left == right == _FLOW else _ANY     # a ratio of flows
        return _ANY
    if isinstance(node, FuncCall):
        if node.func in ('IF', 'CHOOSE', 'MIN', 'MAX'):
            return _joined(kinds)
        if node.func in ('ROUND', 'INT', 'ABS', 'MOD'):
            return kinds[0] if kinds else _CONST
        if node.func in ('SUM', 'AVERAGE') and _per_period_call(node):
            return _joined(kinds)
        if node.func == 'PRODUCT' and _per_period_call(node):
            if _is_flow_product(node):
                factors = [a for a in node.args if isinstance(a, Expr)]
                return _product_kind(node, factors, kinds, memo)
            return _joined(kinds)
        if node.func in _DAY_COUNTS:
            return _FLOW            # a count of the period's days: its length
        return _ANY
    if isinstance(node, (MethodCall, Restrict)) and kinds:
        return kinds[0]
    if isinstance(node, Regrain):
        return _FLOW if node.recipe == 'sum' else _LEVEL
    return _ANY


def _expr_kind(expr: Any, memo: Dict[Any, Any]) -> str:
    """How ``expr``'s values add up over a coarser period — ``_FLOW``,
    ``_LEVEL``, ``_CONST``, ``_SCHEDULE`` or ``_ANY``: a row with a rule by
    it, a formula with none by what it is made of (a ruled flow less a
    ruled flow is a flow), a product of two lines as a flow, a number times
    a value or a value over a number as that value, a quotient of two lines
    as a level, a condition or a cap as its branches when they agree. Worked
    out with a stack, through any number of lines; rows that read each
    other with no rule (a loop) have no kind."""
    rows: Dict[int, str] = memo.setdefault('kinds', {})
    kinds: Dict[int, str] = {}
    open_rows: set = set()
    looped: set = set()
    stack: List[Tuple[Any, bool]] = [(expr, False)]
    while stack:
        node, ready = stack.pop()
        if id(node) in kinds:
            continue
        row = node.var if isinstance(node, VarRef) else None
        if row is not None and not ready:
            if id(row) in rows:
                kinds[id(node)] = rows[id(row)]
                continue
            if id(row) in open_rows:
                kinds[id(node)] = _ANY
                looped.add(id(node))
                continue
        parts = _kind_parts(node, memo)
        if not ready and parts:
            if row is not None:
                open_rows.add(id(row))
            stack.append((node, True))
            stack.extend((part, False) for part in parts if id(part) not in kinds)
            continue
        kind = _node_kind(node, [kinds[id(part)] for part in parts], memo)
        kinds[id(node)] = kind
        if any(id(part) in looped for part in parts):
            looped.add(id(node))
        if row is not None:
            open_rows.discard(id(row))
            if id(node) not in looped:
                rows[id(row)] = kind
    return kinds[id(expr)]


def _turns_on_position(node: Any, memo: Dict[Any, Any]) -> Any:
    """What ``node`` reads of the period's number or dates, at any distance:
    True for a time row of them, ``_ROLLED`` for a running total built
    inline on them or on the period's length, the row for a line of the
    model whose value turns on them (a schedule, or a line that mixes one
    with stored data), None for none."""
    from .expr import TimeRef
    quiet: set = memo.setdefault('quiet', set())     # inline series that read none
    seen: set = set()
    stack = [node]
    while stack:
        item = stack.pop()
        while isinstance(item, Paren):
            item = item.inner
        if isinstance(item, TimeRef):
            if item.field in _POSITION_FIELDS:
                return True
            continue
        if isinstance(item, VarRef):
            var = item.var
            field = _time_row_field(var)
            if field is not None:
                if field in _POSITION_FIELDS:
                    return True
                continue
            sub = _inline_series(var)
            if sub is not None and not isinstance(sub, SelfRef):
                if id(var) not in seen and id(var) not in quiet:
                    seen.add(id(var))
                    stack.append(sub)
                continue
            if _turns_on_dates(var, memo):
                return var if getattr(var, '_owner', None) is not None else _ROLLED
            continue
        if not isinstance(item, Expr) or isinstance(item, Subscript):
            continue
        if isinstance(item, BinOp) and item.op in ('+', '-'):
            stack.extend(_beyond_duration(_signed_terms(item)))
            continue
        stack.extend(_parts(item))
    # Every inline series walked here reads none (a total of the revenues
    # shared by a thousand lines is walked once).
    quiet.update(seen)
    return None


def _condition_name(node: Any) -> Optional[str]:
    if isinstance(node, Compare):
        return 'a comparison'
    if isinstance(node, FuncCall) and node.func in _RECOMPUTE_FUNCS:
        return f"mo.{node.func}"
    return None


def _position_cause(expr: Any, memo: Dict[Any, Any],
                    conditions: bool = True) -> Optional[Tuple[Any, Any]]:
    """Why a formula with no re-grain rule has no total of its own, as
    ``(row, operation)``: arithmetic on the period's number or dates
    ``(None, None)``; a condition, a cap, a power, a quotient or a product
    of schedules on them ``(None, 'mo.IF')``, or on a line whose value turns
    on them ``(row, 'mo.IF')`` — unless the condition yields what may be a
    flow (a tax on a revenue that has its rule), which sums as 0.2.3 summed
    it; it is refused when it yields fixed numbers, a level or a schedule. A
    product of two lines that vary by period that adds up as a flow (stored
    data in it, ``units * price``; the period's length, ``tariff * flag *
    days``; a flow among its factors, ``gen * tariff``) is computed period
    by period and summed, whatever its factors read; a duration on the
    dates and the period's length are no position; a
    row read linearly answers for itself. With ``conditions`` False (the
    source of a running total, whose sum is all it needs) a condition is no
    cause. None when nothing is."""
    from .expr import TimeRef
    clean: set = memo.setdefault('clean', set())     # (inline series, region, conditions) with no cause
    seen: set = set()
    stack: List[Tuple[Any, str]] = [(expr, _RECOMPUTED)]
    while stack:
        node, region = stack.pop()
        while isinstance(node, Paren):
            node = node.inner
        where = 'a power' if region == _POWER else None
        if isinstance(node, TimeRef):
            if node.field in _POSITION_FIELDS:
                return None, where
            continue
        if isinstance(node, VarRef):
            var = node.var
            field = _time_row_field(var)
            if field is not None:
                if field in _POSITION_FIELDS:
                    return None, where
                continue
            sub = _inline_series(var)
            if sub is not None and not isinstance(sub, SelfRef):
                if (id(var), region) not in seen and (id(var), region, conditions) not in clean:
                    seen.add((id(var), region))
                    stack.append((sub, region))
                continue
            if sub is not None and (id(var), region) not in seen:
                # a running total built inline: its source needs only its sum
                seen.add((id(var), region))
                inner = _inline_series(getattr(var, '_cumsum_source', None))
                cause = None if inner is None else _position_cause(inner, memo, conditions=False)
                if cause is not None:
                    return cause
            if region == _POWER and _turns_on_dates(var, memo):
                return (var if getattr(var, '_owner', None) is not None else _ROLLED), where
            continue
        if not isinstance(node, Expr) or isinstance(node, Subscript):
            continue
        if isinstance(node, BinOp) and node.op in ('+', '-'):
            stack.extend((term, region) for term in _beyond_duration(_signed_terms(node)))
            continue
        if _is_flow_product(node):
            if region == _RECOMPUTED and (not conditions or _expr_kind(node, memo) == _FLOW):
                continue                # a flow, summed period by period
            reads = _turns_on_position(node, memo)      # a product of schedules: a schedule
            if reads is not None:
                return (None if reads is True else reads), where or 'a product'
            continue
        name = _condition_name(node)
        if name is not None:
            if region == _RECOMPUTED and (
                    not conditions or _expr_kind(node, memo) in (_FLOW, _ANY)):
                continue                # summed period by period: a flow, or as 0.2.3 did
            reads = _turns_on_position(node, memo)
            if reads is not None:
                return (None if reads is True else reads), where or name
            continue
        if isinstance(node, BinOp) and node.op == '**':
            stack.extend(((node.left, _POWER), (node.right, _POWER)))
            continue
        if (isinstance(node, BinOp) and node.op == '/' and region == _RECOMPUTED
                and (_expr_kind(node.left, memo) in (_CONST, _LEVEL) or _is_gate(node.left))):
            # A number or a level over a line of the dates (loan / months
            # left): the quotient of totals is no total. A flow over one is
            # a rate, recomputed as a ratio of totals.
            reads = _turns_on_position(node.right, memo)
            if reads is not None:
                return (None if reads is True else reads), 'a quotient'
        stack.extend((part, region) for part in _sub_exprs(node))
    # No cause anywhere: every inline series walked here has none either.
    clean.update((key, region, conditions) for key, region in seen)
    return None


def _holds_dates(value: Any) -> bool:
    from datetime import date

    from .tracks import TrackValues
    tracks = list(value.values()) if isinstance(value, TrackValues) else [value]
    return any(isinstance(item, date)
               for track in tracks for item in (track if isinstance(track, list) else [track]))


def _is_time_arithmetic(var: Any) -> bool:
    """A line of dates read off the period's own dates and fixed numbers
    (``mo.time.end + 1``), which re-projects as time — not dates counted by
    the period's number (``close + 30 * mo.time.index``) nor dates moved by
    a condition (``mo.time.end + mo.IF(…)``)."""
    if not _holds_dates(getattr(var, '_value', None)):
        return False
    for _sign, term in _signed_terms(var._expr):
        if _date_of(term) is not None or isinstance(term, Literal):
            continue
        if isinstance(term, VarRef) and not _is_series(getattr(term.var, '_value', None)):
            continue                # a date or a number of the inputs
        return False
    return True


def _own_position_cause(var: Any, grain: str, cache: Dict[int, Any]) -> Optional[Tuple[Any, Any]]:
    """Why ``var`` — a named formula with no re-grain rule — has no ``grain``
    total of its own, or None. A line already at ``grain`` needs none; an
    inline part of a line is judged with its line; a row that IS time, and
    a line of time arithmetic, re-project as time."""
    if getattr(var, '_owner', None) is None or _time_row_field(var) is not None:
        return None
    loc = var.time
    if loc is not None and loc.grain == grain:
        return None
    expr = getattr(var, '_expr', None)
    if expr is None or getattr(var, '_regrain', None) is not None:
        return None
    memo = cache.setdefault(id(_POSITION_MEMO), {})
    verdicts = memo.setdefault('verdicts', {})
    if id(var) not in verdicts:
        verdicts[id(var)] = None if _is_time_arithmetic(var) else _position_cause(expr, memo)
    return verdicts[id(var)]


def _own_refusal(var: Any, cause: Tuple[Any, Any], grain: str) -> ValueError:
    row, where = cause
    label = _line_label(var)
    if row is None and where is None:
        head = f"{label!r} depends on {_DATES}"
    elif row is None:
        head = f"{label!r} puts {_DATES} through {where}"
    elif row is _ROLLED:
        head = (f"{label!r} puts a running total built on the period's number, "
                f"dates or length through {where}")
    else:
        label, name = _labels(var, row)
        head = f"{label!r} puts {name!r}, which depends on {_DATES}, through {where}"
    return _position_refusal(
        f"{head}, so its {grain} total is not defined without a re-grain rule: "
        f"{_RULES}. Lines computed from it with a number do not follow its rule "
        f"safely — capacity * ramp totals as a level when ramp is one, 1 - flag "
        f"counts the 1 once — so give such a line a rule of its own.",
        line=var, source=var)


def _refuse_position_without_rule(var: Variable, grain: str, cache: Dict[int, Any]) -> None:
    """A named formula with no re-grain rule whose total turns on the
    period's number or dates has no total of its own: recomputed it gives
    the quarter's or the year's number (a payroll grown by
    ``mo.time.index`` totalled one month's pay), summed it turns a level
    into a flow (a price of 50 switched by date totalled 600 a year)."""
    cause = _own_position_cause(var, grain, cache)
    if cause is not None:
        raise _own_refusal(var, cause, grain)


def _refuse_rolled_position(var: Any, grain: str, cache: Dict[int, Any]) -> None:
    """A running total of a formula built inline re-grains as the running
    total of that formula's re-grained sums: right for any condition, wrong
    for arithmetic on the period's number or dates (``mo.cumsum(k *
    mo.time.index)``), which is refused."""
    source = var._cumsum_source
    sub = _inline_series(source)
    if getattr(var, '_owner', None) is None or sub is None:
        return
    loc = var.time
    if loc is not None and loc.grain == grain:
        return
    cause = _position_cause(sub, cache.setdefault(id(_POSITION_MEMO), {}), conditions=False)
    if cause is not None:
        raise _own_refusal(var, cause, grain)


def _computed_from(var: Any, error: ValueError, grain: str,
                   cache: Dict[int, Any]) -> ValueError:
    """The refusal of a line met while working out one refused for the
    period's number or dates: its own reason when it has one, else that it
    is computed from that line, named, with the rule that holds."""
    cause = _own_position_cause(var, grain, cache)
    if cause is not None:
        return _own_refusal(var, cause, grain)
    source = error.__dict__['position_source']
    label, name = _labels(var, source)
    return _position_refusal(
        f"{label!r} is computed from {name!r}, which depends on {_DATES}, so its "
        f"{grain} total is not defined until {name!r} has a re-grain rule "
        f"({_RULES}); {label!r} then totals from the total of {name!r}. If "
        f"{label!r} is a flow computed from a level with a number (capacity * "
        f"ramp), or adds a number to a flow (1 - flag), give it a rule of its "
        f"own too.", line=var, source=source)


def _refuse_instead(var: Any, grain: str, cache: Dict[int, Any], exc: Exception) -> None:
    """Where the recompute of a line's formula cannot re-run it (a date
    function, dates that do not add up) and the line's total turns on the
    period's number or dates, the refusal says that, and the rule — raised
    in place of ``exc``. Called on the error path only, so a long chain of
    formulas keeps every frame it had."""
    if isinstance(exc, ValueError) and not str(exc).startswith(_UNSUPPORTED):
        return
    cause = _own_position_cause(var, grain, cache)
    if cause is not None:
        raise _own_refusal(var, cause, grain) from None


_UNSUPPORTED = "re-grain recompute does not support"


def _unsupported(what: str) -> str:
    return (f"{_UNSUPPORTED} {what}; give this formula "
            f"an explicit rule (e.g. .set_regrain(mo.up('sum'))).")


def _recompute(expr: Any, grain: str, cache: Dict[int, Any]) -> Any:
    """Re-run a formula's expression with each dependency re-grained, using the
    same operators (chokepoint) — not a parallel evaluator. Products of two
    time-varying subexpressions divert to evaluate-then-regrain (see
    :func:`_flow_regrain`)."""
    if isinstance(expr, Literal):
        return expr.value
    if isinstance(expr, VarRef):
        return project_variable(expr.var, grain, cache)
    if isinstance(expr, Paren):
        return _recompute(expr.inner, grain, cache)
    if isinstance(expr, UnaryOp):
        fn = _UNARYOPS.get(expr.op)
        if fn is None:
            raise ValueError(_unsupported(f"unary {expr.op!r}"))
        return fn(_recompute(expr.operand, grain, cache))
    if isinstance(expr, BinOp):
        if (expr.op == '*'
                and _is_time_varying(expr.left) and _is_time_varying(expr.right)):
            return _flow_regrain(expr, grain)
        fn = _BINOPS.get(expr.op)
        if fn is None:
            raise ValueError(_unsupported(f"operator {expr.op!r}"))
        return fn(_recompute(expr.left, grain, cache),
                  _recompute(expr.right, grain, cache))
    if (isinstance(expr, FuncCall) and expr.func in _PER_PERIOD_AGGREGATES
            and _per_period_call(expr)):
        if expr.func == 'PRODUCT' and sum(map(_is_time_varying, expr.args)) >= 2:
            return _flow_regrain(expr, grain)
        from .. import functions as fns
        return getattr(fns, expr.func)(*[_recompute(a, grain, cache) for a in expr.args])
    if isinstance(expr, (Compare, FuncCall)) and (
            not isinstance(expr, FuncCall) or expr.func in _RECOMPUTE_FUNCS):
        # NON-LINEAR forms (mo.IF — the tax line — comparisons,
        # MIN/MAX caps) don't commute with aggregation: recomputing
        # IF over quarterly EBIT breaks the tie to the monthly cash
        # tax the model actually pays. Same law as products of two
        # time-varying lines: evaluate at NATIVE grain, re-grain the
        # output by sum (flow default; declare a rule to override).
        return _flow_regrain(expr, grain)
    if (isinstance(expr, MethodCall) and expr.method == 'shift'
            and isinstance(expr.base, VarRef)):
        # ``mo.lag`` — shift IN THE TARGET GRAIN: the re-grained base
        # shifted by the same period count. Both lag idioms stay correct
        # in the coarse grain: ΔNWC_q = nwc_q − nwc_{q−1};
        # opening_q = closing_{q−1}.
        from ..functions.lag import lag as _lag
        base = project_variable(expr.base.var, grain, cache)
        if not isinstance(base, Variable) or not isinstance(base._value, list):
            raise ValueError(_unsupported("a lag over a non-series input"))
        periods = expr.args[0] if expr.args else 1
        fill_node = expr.kwargs.get('fill_value', 0.0)
        if isinstance(fill_node, VarRef):
            fill: Any = project_variable(fill_node.var, grain, cache)
        elif isinstance(fill_node, Expr):
            raise ValueError(_unsupported("a non-literal lag fill"))
        else:
            fill = fill_node
        result = _lag(base, int(periods), fill=fill)
        return result
    if (isinstance(expr, MethodCall) and expr.method == 'copy'
            and not expr.args and not expr.kwargs
            and isinstance(expr.base, VarRef)):
        # An ALIAS — ``report.revenue = pnl.revenue`` — which adoption
        # into a second parent renders as ``revenue.copy()``. Nothing
        # is computed, so there is nothing to re-derive: project the
        # source and hand back its copy. Without this, the bread-and-
        # butter idiom of surfacing one line in another block refuses
        # to change grain even though the source projects cleanly.
        base = project_variable(expr.base.var, grain, cache)
        if not isinstance(base, Variable):
            raise ValueError(_unsupported("an alias of a non-Variable"))
        return base.copy()
    if isinstance(expr, SelfRef):
        # A recurrence's formula is self-feeding — re-running it at a
        # coarser grain would compound at the wrong frequency. Only its
        # VALUES can roll up, and stocks/flows/rates roll differently,
        # so the author must say which.
        raise ValueError(
            "a recurrence's formula is grain-frozen — declare how its "
            "VALUES re-grain: .set_regrain(mo.frozen(values='last')) "
            "for balances / roll-forwards, values='sum' for flows "
            "(costs, one-time capex), values='mean' for prices / rates."
        )
    raise ValueError(_unsupported(f"{type(expr).__name__} expressions"))


def _per_period_call(call: FuncCall) -> bool:
    """A call over as many operands as its ``per_period_arity`` or more, a
    series among them — one call per period."""
    from ..functions._helpers import series_node
    from ..functions.registry import per_period_arity
    arity = per_period_arity(call.func)
    return (arity is not None and len(call.args) >= arity
            and any(series_node(arg) for arg in call.args))


def _is_time_varying(expr: Any, _seen: Optional[set] = None) -> bool:
    """Does this subexpression reference any time-located Variable?

    Looks THROUGH a floating list intermediate: ``a * r`` is a series
    with no owner and so no ``.time`` of its own, yet it varies exactly
    as ``a`` does — taken for a constant, ``(a * r) * g`` re-grained as
    Σa·r·Σg instead of Σ(a·r·g)."""
    if isinstance(expr, Literal):
        return False
    if isinstance(expr, VarRef):
        var = expr.var
        if var.time is not None:
            return True
        sub = getattr(var, '_expr', None)
        if sub is None or not isinstance(var._value, list):
            return False
        seen = _seen if _seen is not None else set()
        if id(var) in seen:
            return False
        seen.add(id(var))
        return _is_time_varying(sub, seen)
    if isinstance(expr, Paren):
        return _is_time_varying(expr.inner, _seen)
    if isinstance(expr, UnaryOp):
        return _is_time_varying(expr.operand, _seen)
    if isinstance(expr, BinOp):
        return (_is_time_varying(expr.left, _seen)
                or _is_time_varying(expr.right, _seen))
    return True  # unknown node: assume time-varying (conservative)


def _first_location(expr: Any, _seen: Optional[set] = None) -> Any:
    """The native ``(start, grain)`` of the first list-valued located ref —
    the evaluation window of a subexpression (all refs in one formula share
    the ambient window, so first-found is representative).

    Walks TRANSITIVELY: a floating intermediate (``ebit > 0`` inside
    ``mo.IF``) has no owner and no ``.time`` of its own — its location
    lives one level down, on the real inputs its expr references.
    """
    if _seen is None:
        _seen = set()
    from .tracks import TrackValues
    for ref in expr.iter_refs():
        loc = ref.time
        value = ref._value
        if loc is not None and (isinstance(value, list) or (
                isinstance(value, TrackValues) and value.time_length is not None)):
            return loc
    for ref in expr.iter_refs():
        if id(ref) in _seen:
            continue
        _seen.add(id(ref))
        sub = getattr(ref, '_expr', None)
        if sub is not None:
            loc = _first_location(sub, _seen)
            if loc is not None:
                return loc
    return None


def _native_eval(expr: Any) -> Any:
    """Evaluate a subexpression at its NATIVE grain through the operator
    chokepoint — refs enter as copies, so the result is an ordinary
    (floating) Variable carrying the native value series."""
    if isinstance(expr, Literal):
        return expr.value
    if isinstance(expr, VarRef):
        return expr.var.copy()
    if isinstance(expr, Paren):
        return _native_eval(expr.inner)
    if isinstance(expr, UnaryOp):
        fn = _UNARYOPS.get(expr.op)
        if fn is None:
            raise ValueError(_unsupported(f"unary {expr.op!r}"))
        return fn(_native_eval(expr.operand))
    if isinstance(expr, BinOp):
        fn = _BINOPS.get(expr.op)
        if fn is None:
            raise ValueError(_unsupported(f"operator {expr.op!r}"))
        return fn(_native_eval(expr.left), _native_eval(expr.right))
    if isinstance(expr, Compare):
        fn = _COMPARES.get(expr.op)
        if fn is None:
            raise ValueError(_unsupported(f"comparison {expr.op!r}"))
        return fn(_native_eval(expr.left), _native_eval(expr.right))
    if isinstance(expr, FuncCall) and expr.func in _RECOMPUTE_FUNCS + _PER_PERIOD_AGGREGATES:
        from .. import functions as _fns
        fn = getattr(_fns, expr.func)
        args = [
            _native_eval(a) if isinstance(a, Expr) else a
            for a in expr.args
        ]
        return fn(*args)
    raise ValueError(_unsupported(f"{type(expr).__name__} expressions"))


def _sum_series(values: list, loc: Any, grain: str,
                source: Optional[Variable] = None,
                display_name: Optional[str] = None) -> Variable:
    """Re-grain a native value series by ``sum`` into a coarse Variable,
    carrying a live ``Regrain`` expr when an addressable source exists
    (renders ``=SUM(range)``; a floating source falls back to inlined
    values per the renderer contract)."""
    from .time import bucket_ranges, regrain_series
    from .time import coarse_first_day
    first_day = getattr(loc, 'first_day', None)
    labels, coarse = regrain_series(values, loc.grain, grain, loc.start, 'sum', first_day)
    ranges = bucket_ranges(loc.grain, grain, loc.start, len(values))
    result = Variable(coarse, display_name=display_name,
                      start=labels[0], grain=grain)
    result._first_day = coarse_first_day(loc.start, loc.grain, first_day, grain)
    if source is not None:
        result._set_expr(Regrain(source=source,
                                 buckets=[(lo, hi) for _, lo, hi in ranges],
                                 recipe='sum', fill_values=list(coarse)))
    return result


def _flow_regrain(expr: Any, grain: str) -> Variable:
    """Evaluate-then-regrain: a product of two time-varying lines is itself a
    flow — evaluate it at native grain, then ``sum`` the output per bucket.
    The default for such products: correct numbers with no rule declared."""
    loc = _first_location(expr)
    if loc is None:
        raise ValueError(
            "this product of time-varying lines has no resolvable native "
            "grain — set grain= on an input, or default_grain on an ancestor "
            "model, so its output can re-grain as a flow."
        )
    native = _native_eval(expr)
    if isinstance(native, Variable):
        from .tracks import TrackValues
        if isinstance(native._value, TrackValues):
            # A tracked product (plan × fx): each track is its own flow.
            projected: Dict[str, Any] = {}
            start = None
            for role, track in native._value.items():
                series = _sum_series(track if isinstance(track, list) else [track],
                                     loc, grain)
                projected[role] = series._value
                start = start or series._start
            result = Variable(display_name=native._display_name)
            result._value = TrackValues(projected)
            result._regrained_tracks = True     # the blend folded on its own too
            result.var_type = 'list'
            result._start, result._grain = start, grain
            from .time import coarse_first_day
            result._first_day = coarse_first_day(loc.start, loc.grain,
                                                 getattr(loc, 'first_day', None), grain)
            return result
        values = native._value if isinstance(native._value, list) else [native._value]
        return _sum_series(values, loc, grain, source=native)
    return _sum_series([native], loc, grain)


def _resolve_ratio_source(var: Variable, name: str) -> Any:
    """Find a ``mo.ratio`` source by name, Python-scope style: the owning
    container first, then each ancestor outward, then the model root by
    dotted path (``'assumptions.users'``).

    Ratio sources are routinely NOT siblings — ARPU divides revenue (in the
    P&L) by users (in assumptions) — so name resolution has to leave the
    container the way an ordinary reference does.
    """
    def _dig(node: Any, path: str) -> Any:
        for part in path.split('.'):
            node = getattr(node, part, None)
            if node is None:
                return None
        return node

    node: Any = var._owner
    seen: set = set()
    while node is not None and id(node) not in seen:
        seen.add(id(node))
        found = _dig(node, name)
        if isinstance(found, Variable):
            return found
        node = getattr(node, '_parent', None) or getattr(node, '_owner', None)
    return None


def _ratio_regrain(var: Variable, num_name: str, den_name: str,
                   grain: str) -> Variable:
    """Evaluate a ``mo.ratio(num, den)`` rule: sum each named source series
    over the shared buckets, then divide — sum-over-sum per bucket, rendered
    live as ``=SUM(num_range)/SUM(den_range)``. Division by an empty bucket
    resolves to an error cell through the chokepoint, never a crash.

    A line with tracks divides track by track: each named line's own series
    on that track (the blend among them), a line without tracks the same
    series on every track - as the book's subtotal reads each track's rows."""
    from .tracks import TrackValues
    label = var._display_name or var.id
    tracked = isinstance(var._value, TrackValues)
    sources = {}
    for role, name in (('num', num_name), ('den', den_name)):
        sibling = _resolve_ratio_source(var, name)
        value = getattr(sibling, '_value', None)
        if not isinstance(sibling, Variable) or not (
                isinstance(value, list) or (
                    tracked and isinstance(value, TrackValues)
                    and value.time_length is not None)):
            raise ValueError(
                f"{label!r} re-grains as mo.ratio({num_name!r}, {den_name!r}), "
                f"but {name!r} doesn't resolve to a list-valued line — name it "
                f"as a sibling ('users'), or by path from an enclosing "
                f"container ('assumptions.users')."
            )
        if sibling.time is None:
            raise ValueError(
                f"{label!r}'s ratio source {name!r} has no resolvable native "
                f"grain — set grain= on it or default_grain on an ancestor."
            )
        sources[role] = sibling
    if not tracked:
        num_c = _sum_series(sources['num']._value, sources['num'].time, grain,
                            source=sources['num'])
        den_c = _sum_series(sources['den']._value, sources['den'].time, grain,
                            source=sources['den'])
        result = num_c / den_c
        result._display_name = var._display_name
        result._grain = grain
        result._start = num_c._start
        result._first_day = num_c._first_day
        return result
    projected: Dict[str, Any] = {}
    first: Any = None
    for track in var._value.roles:
        parts = []
        for side in ('num', 'den'):
            src = sources[side]
            series = src._value
            if isinstance(series, TrackValues):
                if track not in series:
                    raise ValueError(
                        f"{label!r} re-grains as mo.ratio({num_name!r}, "
                        f"{den_name!r}) on track {track!r}, which "
                        f"{src._display_name or src.id!r} does not carry."
                    )
                series = series[track]
            if not isinstance(series, list):
                # A constant track holds in every period of the line.
                length = getattr(src._value, 'time_length', None) or len(var._value[track])
                series = [series] * length
            parts.append(_sum_series(list(series), src.time, grain))
        projected[track] = (parts[0] / parts[1])._value
        if first is None:
            first = parts[0]
    result = Variable(display_name=var._display_name)
    result._value = TrackValues(projected)
    result._regrained_tracks = True     # each track - the blend too - divided on its own
    result.var_type = 'list'
    result._unit = var._unit
    result._grain = grain
    result._start = first._start
    result._first_day = first._first_day
    return result


#: Marks a Variable whose projection is in progress. Meeting it again
#: means the walk came back round a loop (rows that read each other
#: across periods, ``core.loops``) — a recompute cannot follow that.
_IN_PROGRESS = object()


def project_variable(var: Variable, grain: str, cache: Dict[int, Any]) -> Any:
    """Re-grain one Variable to ``grain`` (memoised by identity)."""
    if getattr(var, '_awaits', None) or getattr(var, '_forward_of', None) is not None:
        from .loops import pending_message
        raise ValueError(pending_message(var))
    key = id(var)
    refused = cache.get(id(_REFUSED))
    if refused and key in refused:
        raise _position_refusal(*refused[key])
    if key in cache:
        if cache[key] is _IN_PROGRESS:
            from .loops import loop_rows
            loop = [v for v in loop_rows(var) if v._regrain is None]
            names = ", ".join(sorted(repr(v.python_name) for v in loop)) \
                or repr(var.python_name)
            raise ValueError(
                f"{names} form a loop across periods (rows that read each "
                f"other through mo.lag); formulas in a loop cannot be re-run "
                f"at another grain. Declare how each row's VALUES re-grain: "
                f".set_regrain(mo.frozen(values='last')) for a balance, "
                f"values='sum' for a flow."
            )
        return cache[key]
    cache[key] = _IN_PROGRESS
    try:
        return _project_variable(var, grain, cache, key)
    except ValueError as exc:
        if not _is_position_refusal(exc):
            raise
        # A line computed from one with no total says so in its own name,
        # and is kept refused, so the lines after it do not work it out anew.
        error = exc
        if (exc.__dict__['position_line'] is not var
                and getattr(var, '_owner', None) is not None):
            error = _computed_from(var, exc, grain, cache)
        cache.setdefault(id(_REFUSED), {})[key] = (
            str(error), error.__dict__['position_line'], error.__dict__['position_source'])
        if error is exc:
            raise
        raise error from None
    finally:
        # A path that does not memoise leaves the marker; drop it so a
        # later reference recomputes as before.
        if cache.get(key) is _IN_PROGRESS:
            del cache[key]


def _project_variable(var: Variable, grain: str, cache: Dict[int, Any], key: int) -> Any:
    from .tracks import TrackValues
    if isinstance(getattr(var, '_value', None), TrackValues):
        # Track-wise projection: the tracks axis passes through the
        # re-grain; each track re-grains by the LINE's own rule. Pure
        # derived lines recompute over their projected operands (the
        # broadcast re-zips every track, live included); lines with
        # authored data project their STORED tracks — a recompute would
        # silently overwrite entered values with formula values.
        from .regrain import Ratio as _Ratio
        _spec = var._regrain
        _ruled = _spec is not None and not isinstance(_spec.default, _Ratio)
        _loc = var.time
        _rule = (None if _spec is None else
                 _spec.overrides.get((_loc.grain, grain), _spec.default)
                 if _spec.overrides and _loc is not None else _spec.default)
        if isinstance(_rule, _Ratio) and not _spec.frozen:
            # The line's own rule, track by track: sum over sum of the named
            # lines' same-track series - never a recompute of the formula
            # over its operands' own rules (a mean denominator).
            result = (var.copy() if _loc is not None and _loc.grain == grain
                      else _ratio_regrain(var, _rule.num, _rule.den, grain))
        elif getattr(var, '_cumsum_source', None) is not None and not _ruled:
            # A running total of a tracked line: each track's running total
            # of its folded input, as an untracked one re-grains below.
            from ..functions.recurrence import cumsum as _cumsum
            _refuse_rolled_position(var, grain, cache)
            result = _cumsum(project_variable(var._cumsum_source, grain, cache))
            result._display_name = var._display_name
        elif (var._role_kwargs is None and var.is_formula
                and var._expr is not None and not _ruled
                and var._value.time_length is None):
            # One cell per track — a reduction over the horizon (a SUM of
            # a plan/actual line, an NPV) — is the same number at every
            # grain, as the untracked branch below copies it.
            result = var.copy()
        elif (var._role_kwargs is None and var.is_formula
                and var._expr is not None and not _ruled):
            try:
                result = _recompute(var._expr, grain, cache)
            except (ValueError, TypeError) as exc:
                _refuse_instead(var, grain, cache, exc)
                raise
            # After the recompute, so that a recurrence or a loop row
            # reports its own error first.
            _refuse_position_without_rule(var, grain, cache)
            if isinstance(result, Variable):
                result._display_name = var._display_name
                result._grain = grain
        else:
            from .regrain import hole_value
            from .variable import TimeLoc
            line_blank = hole_value(var._value)
            label = var._display_name or var.python_name
            # Each role re-grains as its own series would: at the line's
            # own placement when it has one, else — a list — at the line's
            # ambient window (the owner chain), else time-agnostic. Resolved
            # once for the line, so no stand-in Variable is built per role
            # (each would be projected through Variable.at with two more
            # Variables and a Regrain).
            own = (TimeLoc(var._start, var._grain, var._first_day)
                   if var._grain is not None else None)
            ambient = own if own is not None else var.time
            projected: Dict[str, Any] = {}
            p_start = None
            p_first = None
            for role, track in var._value.items():
                loc_r = own if own is not None else (
                    ambient if isinstance(track, list) else None)
                values, start, first = _fold_track(
                    track, loc_r, _spec, grain, label, line_blank,
                    var._start, var._first_day,
                )
                projected[role] = values
                if p_start is None:
                    p_start = start
                    p_first = first
            result = Variable(display_name=var._display_name)
            result._value = TrackValues(projected)
            # Each track — the blend among them — re-grained on its own:
            # the blend is NOT a splice of the re-grained tracks (a close
            # inside a quarter splits that quarter).
            result._regrained_tracks = True
            result.var_type = (
                'list' if result._value.time_length is not None else 'scalar'
            )
            result._unit = var._unit
            result._grain = grain
            result._start = p_start
            result._first_day = p_first
        if (isinstance(result, Variable)
                and not getattr(result, '_excel_props', None)
                and var._excel_props):
            result._excel_props = dict(var._excel_props)
        cache[key] = result
        return result
    if getattr(var, '_indexed_by', ()) or ():
        # Axised gate, rank ≥ 1: an axised Variable's flat value
        # enumerates COORDINATES, not periods — even a single finite
        # axis would have its coordinates summed into quarters with no
        # error. Fiber-wise projection is not supported. Refusing
        # loudly rides the normal absorb path — one error cell, never
        # silently re-bucketed coordinates.
        raise ValueError(
            "an axised Variable cannot change grain — project a "
            "slice (drop the axis first), or keep the model at native "
            "grain."
        )
    loc = var.time
    from .expr import TimeRef
    if isinstance(getattr(var, '_expr', None), TimeRef) and loc is not None:
        # Time is recomputed at the target grain — the days of a quarter,
        # its first and last day — never rolled up from months.
        result: Any = _project_time_row(var, loc, grain)
    elif loc is not None and loc.grain == grain:
        result = var.copy()                             # already at the target grain
    elif var._regrain is not None:
        from .regrain import Ratio
        spec = var._regrain
        rule: Any = spec.default
        if spec.overrides and loc is not None:
            rule = spec.overrides.get((loc.grain, grain), spec.default)
        if isinstance(rule, Ratio) and not spec.frozen:
            result = _ratio_regrain(var, rule.num, rule.den, grain)
        else:
            result = var.at(grain)                      # explicit rule -> apply-own-rule
    elif (getattr(var, '_cumsum_source', None) is not None
            and isinstance(var._value, list)):
        # A running total re-grains exactly as the cumsum of its
        # re-grained (summed) input: cumsum(monthly)[quarter-end]
        # == cumsum(quarterly sums)[quarter].
        from ..functions.recurrence import cumsum as _cumsum
        _refuse_rolled_position(var, grain, cache)
        base = project_variable(var._cumsum_source, grain, cache)
        result = _cumsum(base)
        result._display_name = var._display_name
    elif var.is_formula and var._expr is not None:
        if not isinstance(var._value, list):
            # A SCALAR-valued formula is grain-independent — a rate
            # conversion ((1 + annual) ** (1/12) - 1) or a reduction
            # over the horizon (NPV, IRR, SUM) is THE SAME NUMBER at
            # every grain. Recomputing it over re-grained inputs
            # is at best wasted work and at worst wrong (NPV's period
            # rate no longer matches quarterly flows). Copy.
            result = var.copy()
        else:
            # A LIST formula re-grains by recomputing over its re-grained
            # inputs, even when it is a floating intermediate (no owner, so
            # .time is None) — its grain comes from its inputs, not from
            # itself. Products of two time-varying lines divert inside
            # _recompute to evaluate-then-regrain. One whose own arithmetic
            # reads the period's number or dates has no default: its
            # re-grain rule decides — checked after the recompute, so that a
            # recurrence or a loop row reports its own error first.
            try:
                result = _recompute(var._expr, grain, cache)
            except (ValueError, TypeError) as exc:
                _refuse_instead(var, grain, cache, exc)
                raise
            _refuse_position_without_rule(var, grain, cache)
            if isinstance(result, Variable):
                result._display_name = var._display_name
                result._grain = grain
    elif loc is None:
        result = var.copy()                             # time-agnostic constant
    else:
        label = var._display_name or var.id
        raise ValueError(
            f"{label!r} has no re-grain rule and is not a formula; give it a "
            f"rule (regrain=mo.up('sum'/'last'/'mean'))."
        )
    # Presentation identity survives the projection: the projected row keeps
    # its unit (a quarterly sum of $ is $) and its excel_props (the
    # article number, header-row placement, per-line formats) so the
    # projected sheet lays out exactly like the native one.
    if isinstance(result, Variable):
        if getattr(var, '_unit', None) is not None and (
                getattr(result, '_unit', None) is None
                or getattr(var, '_unit_declared', False)):
            # The unit the author declared is the row's at every grain (a
            # pure number's label does not survive the recomputed product).
            result._unit = var._unit
            result._unit_declared = getattr(var, '_unit_declared', False)
        if not getattr(result, '_excel_props', None) and var._excel_props:
            result._excel_props = dict(var._excel_props)
    cache[key] = result
    return result


def _fold_track(track: Any, loc: Any, spec: Any, grain: str, label: str, blank: Any,
                start: Any, first_day: Any) -> tuple:
    """One role of a tracked line at ``grain`` — ``(values, start, first_day)``,
    what projecting the role as its own series computed: a scalar broadcasts,
    a series already at the grain or without a time copies, a located series
    applies the LINE's rule (:func:`~.variable.series_at_grain`, so the
    refusals name the line the way ``Variable.at`` does)."""
    if not isinstance(track, list):
        return track, start, first_day
    if loc is not None and loc.grain == grain:
        return list(track), start, first_day
    if spec is not None:
        if loc is None:
            if not spec.frozen:
                raise ValueError(
                    f"{label!r} has a re-grain rule but no resolvable native "
                    f"grain — set grain= on it, or default_grain on an ancestor "
                    f"model, so the rule knows which grain it re-grains from."
                )
            return list(track), start, first_day
        from .variable import series_at_grain
        values, start_label, first, _ranges, _rule = series_at_grain(
            label, track, loc, spec, grain, blank)
        return values, start_label, first
    if loc is None:
        return list(track), start, first_day
    raise ValueError(
        f"{label!r} has no re-grain rule and is not a formula; give it a "
        f"rule (regrain=mo.up('sum'/'last'/'mean'))."
    )


def _absorbed(child: Variable, exc: BaseException) -> Variable:
    """The ``'#VALUE!'`` row a line that cannot re-grain becomes under
    ``on_error='absorb'`` — its teaching message on ``_regrain_error``, the
    tracks SHAPE kept (the axis is structure, the token is the value), the
    source text carried so a renderer still shows the row's code."""
    from .tracks import TrackValues as _TV
    projected = Variable('#VALUE!', display_name=child._display_name)
    src_val = getattr(child, '_value', None)
    if isinstance(src_val, _TV):
        projected._value = _TV({r: '#VALUE!' for r in src_val.roles})
        # The blend among them is the error too, never a track to splice
        # again (nor one the row authored).
        projected._regrained_tracks = True
    projected._regrain_error = str(exc)
    src_code = getattr(child, '_source_code', None)
    if src_code:
        projected._source_code = src_code
    return projected


def project_line(var: Variable, grain: str, cache: Dict[int, Any],
                 on_error: str = 'raise') -> Variable:
    """One line of a tree at ``grain`` through the shared ``cache`` — the
    step :func:`project_model` takes per line, available on its own to
    project one line without projecting the whole tree.
    ``on_error='absorb'`` turns a refusal into
    the same ``'#VALUE!'`` row the tree projection would place."""
    try:
        return project_variable(var, grain, cache)
    except (ValueError, TypeError) as exc:
        if on_error != 'absorb':
            raise
        return _absorbed(var, exc)


class _WindowProbe:
    """A bare object for :func:`window_projects` to set a window on."""


def window_projects(mv: Any, grain: str) -> bool:
    """Whether ``mv``'s own window re-grains to ``grain`` — the refusal
    :func:`project_model` raises for a ROOT before touching a line (a model
    of years asked for quarters), answered without a projection."""
    try:
        _propagate_window(mv, _WindowProbe(), grain)
    except Exception:  # noqa: BLE001 — any refusal: the window does not project
        return False
    return True


def project_model(mv: Any, grain: str, cache: Optional[Dict[int, Any]] = None,
                 on_error: str = 'raise') -> Any:
    """Re-grain a whole MV tree onto ``grain`` — returns a new MultiVariable.

    ``on_error='absorb'`` turns a line that cannot re-grain into an
    ``'#VALUE!'`` error cell carrying its teaching message on
    ``_regrain_error`` — one bad line is one error row, never a refused
    projection. ``'raise'`` (default) keeps the strict behavior.
    """
    from .multi_variable import MultiVariable, MultiVariableBase
    is_root = cache is None
    if cache is None:
        cache = {}
    result = MultiVariable(mv._display_name, excel_props=dict(mv._excel_props) or None)
    cache[id(mv)] = result
    for name, child in mv._components.items():
        if isinstance(child, MultiVariableBase):
            from ..compile.excel.view import ExcelView

            if isinstance(child, ExcelView):
                # A named view is presentation metadata — its config
                # leaves carry no time axis and must not be re-grained
                # (a 'quarter' string has no regrain rule). Carry the
                # view over untouched.
                setattr(result, name, child)
                continue
            setattr(result, name, project_model(child, grain, cache, on_error))
        else:
            setattr(result, name, project_line(child, grain, cache, on_error))
    try:
        _propagate_window(mv, result, grain)
    except ValueError:
        if on_error != 'absorb' or is_root:
            raise
        # A container on a coarser grain than the one asked (a sheet of
        # years in a quarterly view) keeps no window of its own: its lines
        # are already the error rows that say why, and the rest of the
        # tree projects. The tree asked as a whole still refuses.
    # The tracks declaration (and its blend) is ambient context exactly
    # like the window — a projected tree that lost it would pick the
    # wrong display default and mis-expand.
    decl = getattr(mv, 'tracks', None)
    if decl is not None:
        result.tracks = decl
    # Keep the root's identity: same python_name -> same crystallized qpaths,
    # so a projected tree keys its lines identically to the native tree.
    if mv._python_name is not None:
        result._python_name = mv._python_name
    _dev = mv.__dict__.get('default_excel_view')
    if _dev is not None:
        # The view cascade rides the projection — a projected tree styles
        # exactly like its source (formats, bands, nesting, meta).
        result.default_excel_view = _dev
    remaps = cache.setdefault('header_remaps', [])      # type: ignore[call-overload]
    if 'default_header' in mv.__dict__:
        remaps.append((result, mv.__dict__['default_header']))
    if is_root:
        _remap_header_slots(remaps, cache)
    return result


def _project_time_row(var: Any, loc: Any, grain: str) -> Variable:
    """``mo.time.<field>`` over the target grain's periods covering
    ``var``'s window: full calendar periods, the same buckets every other
    line of the projection is grouped into."""
    from .time import (DefaultWindow, bucket_ranges, time_field_series,
                       time_variable)
    from .time import coarse_first_day
    field = var._expr.field
    n = len(var._value) if isinstance(var._value, list) else 1
    first_day = getattr(loc, 'first_day', None)
    if loc.grain == grain:
        return time_variable(field, DefaultWindow(loc.start, grain, n, None, first_day))
    buckets = bucket_ranges(loc.grain, grain, loc.start, n)
    # The first coarse period starts where the native window does: short
    # when that is inside it — like every line re-grained beside it.
    coarse_first = coarse_first_day(loc.start, loc.grain, first_day, grain)
    full = time_field_series('days', buckets[0][0], grain, len(buckets), coarse_first)
    native_days = time_field_series('days', loc.start, loc.grain, n, first_day)
    def whole(k: int, lo: int, hi: int) -> bool:
        got = sum(native_days[lo:hi])
        # An hour's 1/24 of a day sums in floating point.
        return got == full[k] if isinstance(got, int) else math.isclose(got, full[k])
    # A clock ruler's periods end where they stop; a day's end is its last
    # day. Over hours, a day's time is its hours' — first start, last
    # stop, their days — never re-spelled at the day's own reading.
    from .calendar import GREGORIAN
    clock = GREGORIAN.knows(loc.grain) and GREGORIAN.ruler(loc.grain).sub_day
    if not clock and all(whole(k, lo, hi) for k, (_, lo, hi) in enumerate(buckets)):
        return time_variable(field, DefaultWindow(buckets[0][0], grain, len(buckets),
                                                  None, coarse_first))
    # A window that starts or ends inside a quarter: its first or last
    # quarter is only the months the window covers — the same months every
    # flow is summed over — so time is those months, as numbers (no
    # calendar formula spells a partial quarter).
    native = {f: time_field_series(f, loc.start, loc.grain, n, first_day)
              for f in ('start', 'end', 'days', 'months')}
    values = []
    for k, (_, lo, hi) in enumerate(buckets):
        values.append({'start': native['start'][lo], 'end': native['end'][hi - 1],
                       'days': sum(native['days'][lo:hi]),
                       'months': sum(native['months'][lo:hi]), 'index': k + 1}[field])
    kind = 'datetime' if field in ('start', 'end') else (
        'int' if all(isinstance(v, int) for v in values) else 'float')
    out = Variable(values, value_type=kind,
                   var_type='list', start=buckets[0][0], grain=grain)
    out._first_day = coarse_first
    recipe = {'start': 'first', 'end': 'last', 'days': 'sum'}.get(field)
    if clock and recipe is not None:
        # Over hours a day's time is read off its hours' own cells — the
        # first start, the last stop, the sum of the days — a live formula
        # wherever the row is laid out, its values where it is not.
        out._set_expr(Regrain(source=var, buckets=[(lo, hi) for _, lo, hi in buckets],
                              recipe=recipe, fill_values=list(values)))
    return out


def _remap_header_slots(remaps: list, cache: Dict[int, Any]) -> None:
    """Point a projected tree's ``default_header`` at the PROJECTED header
    rows — the source's would put monthly rows over quarterly sheets. A
    header the projection did not build (outside the projected subtree)
    is dropped: the projected book then shows its time without one."""
    for dst, header in remaps:
        if header is None:
            dst.default_header = None
            continue
        if isinstance(header, list):
            rows = [cache.get(id(v)) for v in header]
            if rows and all(isinstance(v, Variable) for v in rows):
                dst.default_header = rows
            continue
        projected = cache.get(id(header))
        if projected is not None:
            dst.default_header = projected


def _propagate_window(src: Any, dst: Any, grain: str) -> None:
    """If ``src`` declares its own projection window (``default_grain`` etc.),
    set the re-grained window on ``dst`` so the projected tab resolves a window
    at the new grain — e.g. a monthly model projected to quarter gets
    ``default_grain='quarter'`` and a quarterly ``default_start``, which is what
    lets its Excel timeline header read ``Q1 2024`` instead of nothing.
    """
    src_grain = getattr(src, 'default_grain', None)
    if src_grain is None:
        return
    src_start = getattr(src, 'default_start', None)
    src_periods = getattr(src, 'default_periods', None)
    # A short first period rides along: the coarser grain's first period
    # starts on the same day (a stub lies inside a quarter, so inside its
    # year too).
    src_first = src.__dict__.get('_default_first_day') if hasattr(src, '__dict__') else None
    if src_grain == grain:
        # Identity projection — the window carries over verbatim.
        # (``bucket_ranges`` coarsens only; same-grain must not reach it.)
        dst.default_grain = src_grain
        if src_start is not None:
            dst.default_start = src_start
        if src_periods is not None:
            dst.default_periods = src_periods
        _carry_first_day(src, dst, src_first)
        return
    from .time import _parse, _target_label, bucket_ranges, coarse_first_day
    # Everything is worked out before ``dst`` changes: a window that
    # cannot re-grain leaves none half set.
    start = periods = first = None
    if src_start is not None:
        start = _target_label(_parse(src_start, src_grain), src_grain, grain)
        if src_periods is not None:
            periods = len(bucket_ranges(src_grain, grain, src_start, src_periods))
        first = coarse_first_day(src_start, src_grain, src_first, grain)
    dst.default_grain = grain
    if start is not None:
        dst.default_start = start
        if periods is not None:
            dst.default_periods = periods
        _carry_first_day(src, dst, first)


def _carry_first_day(src: Any, dst: Any, first_day: Any) -> None:
    """Copy a short first period onto a projected window."""
    if first_day is not None:
        dst.__dict__['_default_first_day'] = first_day
