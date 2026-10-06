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
  numbers; declare a rule only to override the flow default.

A shared cache re-grains each source once, so cross-line references agree.
"""

from __future__ import annotations

import operator
from typing import Any, Dict, Optional

from .expr import (
    BinOp, Compare, Expr, FuncCall, Literal, MethodCall, Paren, Regrain,
    SelfRef, UnaryOp, VarRef,
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
#: element-wise or scalar-broadcasting, no self-reference. Reductions
#: (SUM/NPV/IRR/…) never reach recompute: scalar-valued formulas copy.
_RECOMPUTE_FUNCS = ('IF', 'AND', 'OR', 'NOT', 'CHOOSE',
                    'MAX', 'MIN', 'ABS', 'ROUND', 'INT', 'MOD')


def _unsupported(what: str) -> str:
    return (f"re-grain recompute does not support {what}; give this formula "
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
    if isinstance(expr, FuncCall) and expr.func in _RECOMPUTE_FUNCS:
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
            result = _cumsum(project_variable(var._cumsum_source, grain, cache))
            result._display_name = var._display_name
        elif (var._role_kwargs is None and var.is_formula
                and var._expr is not None and not _ruled):
            result = _recompute(var._expr, grain, cache)
            if isinstance(result, Variable):
                result._display_name = var._display_name
                result._grain = grain
        else:
            from .regrain import hole_value
            line_blank = hole_value(var._value)
            projected: Dict[str, Any] = {}
            p_start = None
            p_first = None
            for role, track in var._value.items():
                # Named as the line: a refusal names the line, never this
                # stand-in's floating id.
                s = Variable(display_name=var._display_name or var.python_name)
                s._value = list(track) if isinstance(track, list) else track
                s._hole_value = line_blank      # the whole line's, not one track's
                s.var_type = 'list' if isinstance(track, list) else 'scalar'
                s._regrain = var._regrain
                s._start = var._start
                s._grain = var._grain
                s._first_day = var._first_day
                s._unit = var._unit
                if getattr(var, '_owner', None) is not None:
                    s.__dict__['_owner'] = var._owner
                p = project_variable(s, grain, {})
                projected[role] = p._value
                if p_start is None:
                    p_start = p._start
                    p_first = getattr(p, '_first_day', None)
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
            # _recompute to evaluate-then-regrain.
            result = _recompute(var._expr, grain, cache)
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
            try:
                projected = project_variable(child, grain, cache)
            except (ValueError, TypeError) as exc:
                if on_error != 'absorb':
                    raise
                projected = Variable('#VALUE!', display_name=child._display_name)
                from .tracks import TrackValues as _TV
                src_val = getattr(child, '_value', None)
                if isinstance(src_val, _TV):
                    # Keep the tracks SHAPE under the error: the axis
                    # is structure, the token is the value, so
                    # .at(track=...) still works on the error row.
                    projected._value = _TV(
                        {r: '#VALUE!' for r in src_val.roles}
                    )
                    # The blend among them is the error too, never a track
                    # to splice again (nor one the row authored).
                    projected._regrained_tracks = True
                projected._regrain_error = str(exc)
                # Keep the row's IDENTITY readable: carry the source
                # text over, so any renderer that shows a row's code
                # still shows it for the error row.
                src_code = getattr(child, '_source_code', None)
                if src_code:
                    projected._source_code = src_code
            setattr(result, name, projected)
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
    if all(sum(native_days[lo:hi]) == full[k] for k, (_, lo, hi) in enumerate(buckets)):
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
    out = Variable(values, value_type='datetime' if field in ('start', 'end') else 'int',
                   var_type='list', start=buckets[0][0], grain=grain)
    out._first_day = coarse_first
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
