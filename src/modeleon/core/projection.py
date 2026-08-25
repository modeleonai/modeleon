# SPDX-License-Identifier: Apache-2.0
"""Projection — re-grain a whole model onto a coarser grain.

``model.at(grain)`` walks the tree and re-grains every line:

- a **time-agnostic** constant broadcasts unchanged;
- a line **with a rule** applies it to its own value (``Variable.at`` —
  apply-own-rule); ``mo.ratio(num, den)`` resolves its sibling sources here
  and re-grains as sum-over-sum;
- a **formula without a rule** recomputes over its re-grained inputs — the
  formula's expression is re-run through the *same operators* (the chokepoint,
  ADR-010), which is correct for additive forms (``gross = revenue - cogs``)
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
    return (f"re-grain recompute does not support {what} yet; give this formula "
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


def _is_time_varying(expr: Any) -> bool:
    """Does this subexpression reference any time-located Variable?"""
    if isinstance(expr, Literal):
        return False
    if isinstance(expr, VarRef):
        return expr.var.time is not None
    if isinstance(expr, Paren):
        return _is_time_varying(expr.inner)
    if isinstance(expr, UnaryOp):
        return _is_time_varying(expr.operand)
    if isinstance(expr, BinOp):
        return _is_time_varying(expr.left) or _is_time_varying(expr.right)
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
    for ref in expr.iter_refs():
        loc = ref.time
        if loc is not None and isinstance(ref._value, list):
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
    labels, coarse = regrain_series(values, loc.grain, grain, loc.start, 'sum')
    ranges = bucket_ranges(loc.grain, grain, loc.start, len(values))
    result = Variable(coarse, display_name=display_name,
                      start=labels[0], grain=grain)
    if source is not None:
        result._set_expr(Regrain(source=source,
                                 buckets=[(lo, hi) for _, lo, hi in ranges],
                                 recipe='sum', fill_values=list(coarse)))
    return result


def _flow_regrain(expr: Any, grain: str) -> Variable:
    """Evaluate-then-regrain: a product of two time-varying lines is itself a
    flow — evaluate it at native grain, then ``sum`` the output per bucket.
    The §5.2 default that replaces the old hard error."""
    loc = _first_location(expr)
    if loc is None:
        raise ValueError(
            "this product of time-varying lines has no resolvable native "
            "grain — set grain= on an input, or default_grain on an ancestor "
            "model, so its output can re-grain as a flow."
        )
    native = _native_eval(expr)
    if isinstance(native, Variable):
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
    resolves to an error cell through the chokepoint, never a crash."""
    label = var._display_name or var.id
    sources = {}
    for role, name in (('num', num_name), ('den', den_name)):
        sibling = _resolve_ratio_source(var, name)
        if not isinstance(sibling, Variable) or not isinstance(sibling._value, list):
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
    num_c = _sum_series(sources['num']._value, sources['num'].time, grain,
                        source=sources['num'])
    den_c = _sum_series(sources['den']._value, sources['den'].time, grain,
                        source=sources['den'])
    result = num_c / den_c
    result._display_name = var._display_name
    result._grain = grain
    result._start = num_c._start
    return result


def project_variable(var: Variable, grain: str, cache: Dict[int, Any]) -> Any:
    """Re-grain one Variable to ``grain`` (memoised by identity)."""
    key = id(var)
    if key in cache:
        return cache[key]
    from .tracks import TrackValues
    if isinstance(getattr(var, '_value', None), TrackValues):
        # P2.5 — track-wise projection: the tracks axis passes through
        # the lens; each track re-grains by the LINE's own rule. Pure
        # derived lines recompute over their projected operands (the
        # broadcast re-zips every track, live included); lines with
        # authored data project their STORED tracks — a recompute would
        # silently overwrite pinned facts with formula values.
        from .regrain import Ratio as _Ratio
        _spec = var._regrain
        _ruled = _spec is not None and not isinstance(_spec.default, _Ratio)
        if (var._role_kwargs is None and var.is_formula
                and var._expr is not None and not _ruled):
            result = _recompute(var._expr, grain, cache)
            if isinstance(result, Variable):
                result._display_name = var._display_name
                result._grain = grain
        else:
            from .regrain import Ratio
            spec = var._regrain
            if spec is not None and isinstance(spec.default, Ratio):
                raise ValueError(
                    f"{var._display_name or var.id!r} re-grains as a "
                    f"ratio and carries tracks — project a slice "
                    f"(.at(track='...')); track-wise ratio projection "
                    f"needs per-track sibling sources."
                )
            projected: Dict[str, Any] = {}
            p_start = None
            for role, track in var._value.items():
                s = Variable(display_name=var._display_name)
                s._value = list(track) if isinstance(track, list) else track
                s.var_type = 'list' if isinstance(track, list) else 'scalar'
                s._regrain = var._regrain
                s._start = var._start
                s._grain = var._grain
                s._unit = var._unit
                if getattr(var, '_owner', None) is not None:
                    s.__dict__['_owner'] = var._owner
                p = project_variable(s, grain, {})
                projected[role] = p._value
                if p_start is None:
                    p_start = p._start
            result = Variable(display_name=var._display_name)
            result._value = TrackValues(projected)
            result.var_type = (
                'list' if result._value.time_length is not None else 'scalar'
            )
            result._unit = var._unit
            result._grain = grain
            result._start = p_start
        if (isinstance(result, Variable)
                and not getattr(result, '_excel_props', None)
                and var._excel_props):
            result._excel_props = dict(var._excel_props)
        cache[key] = result
        return result
    if getattr(var, '_indexed_by', ()) or ():
        # §14.4 gate, widened to rank≥1 by the §17 pass: the ONLY v1
        # shape is one finite axis, and it slipped the original rank≥2
        # check — the lens summed coordinates into quarters with no
        # error. Fiber-wise projection is the track layer's job (P2.5).
        # Refusing loudly rides the normal absorb path — one red cell,
        # never silently re-bucketed coordinates.
        raise ValueError(
            "an axised Variable cannot change grain yet — project a "
            "slice (drop the axis first), or keep the model at native "
            "grain."
        )
    loc = var.time
    if loc is not None and loc.grain == grain:
        result: Any = var.copy()                        # already at the target grain
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
            # every view grain. Recomputing it over re-grained inputs
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
    # Presentation identity survives the lens: the projected row keeps
    # its unit (a quarterly sum of ₸ is ₸) and its excel_props (the
    # article number, header-row placement, per-line formats) so the
    # projected sheet lays out exactly like the native one.
    if isinstance(result, Variable):
        if (getattr(result, '_unit', None) is None
                and getattr(var, '_unit', None) is not None):
            result._unit = var._unit
        if not getattr(result, '_excel_props', None) and var._excel_props:
            result._excel_props = dict(var._excel_props)
    cache[key] = result
    return result


def project_model(mv: Any, grain: str, cache: Optional[Dict[int, Any]] = None,
                 on_error: str = 'raise') -> Any:
    """Re-grain a whole MV tree onto ``grain`` — returns a new MultiVariable.

    ``on_error='absorb'`` turns a line that cannot re-grain into an
    ``'#VALUE!'`` error cell carrying its teaching message on
    ``_regrain_error`` — one bad line is one red row, never a refused
    projection. ``'raise'`` (default) keeps the strict behavior.
    """
    from .multi_variable import MultiVariable, MultiVariableBase
    if cache is None:
        cache = {}
    result = MultiVariable(mv._display_name, excel_props=dict(mv._excel_props) or None)
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
            except ValueError as exc:
                if on_error != 'absorb':
                    raise
                projected = Variable('#VALUE!', display_name=child._display_name)
                from .tracks import TrackValues as _TV
                src_val = getattr(child, '_value', None)
                if isinstance(src_val, _TV):
                    # Keep the tracks SHAPE under the error: the axis
                    # is structure, the token is the value. Consumers
                    # keep offering track views while the line teaches
                    # its re-grain rule.
                    projected._value = _TV(
                        {r: '#VALUE!' for r in src_val.roles}
                    )
                projected._regrain_error = str(exc)
                # Keep the row's IDENTITY readable: without the source
                # text, selecting the red cell shows an empty formula
                # bar and the row reads as "code disappeared".
                src_code = getattr(child, '_source_code', None)
                if src_code:
                    projected._source_code = src_code
            setattr(result, name, projected)
    _propagate_window(mv, result, grain)
    # The tracks declaration (and its blend) is ambient context exactly
    # like the window — a projected tree that lost it would pick the
    # wrong display default and mis-expand (§16 P3).
    decl = getattr(mv, 'tracks', None)
    if decl is not None:
        result.tracks = decl
    # Keep the root's identity: same python_name -> same crystallized qpaths,
    # so a projected tree keys its lines identically to the native run.
    if mv._python_name is not None:
        result._python_name = mv._python_name
    _dev = mv.__dict__.get('default_excel_view')
    if _dev is not None:
        # The view cascade rides the lens — a projected tree styles
        # exactly like its source (formats, bands, nesting, meta).
        result.default_excel_view = _dev
    return result


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
    if src_grain == grain:
        # Identity projection — the window carries over verbatim.
        # (``bucket_ranges`` coarsens only; same-grain must not reach it.)
        dst.default_grain = src_grain
        if src_start is not None:
            dst.default_start = src_start
        if src_periods is not None:
            dst.default_periods = src_periods
        return
    from .time import _parse, _target_label, bucket_ranges
    dst.default_grain = grain
    if src_start is not None:
        dst.default_start = _target_label(_parse(src_start, src_grain), src_grain, grain)
        if src_periods is not None:
            dst.default_periods = len(
                bucket_ranges(src_grain, grain, src_start, src_periods)
            )
