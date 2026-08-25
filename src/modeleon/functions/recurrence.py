# SPDX-License-Identifier: Apache-2.0
"""Recurrence helpers — period-over-period series with cell-back-referencing formulas.

- ``recurrence`` — ``x[0] = start; x[t] = formula(prev=x[t-1], ...)`` for t≥1.
  Supports string templates (``"{prev} * (1 + {rate})"``) and lambdas.
  The AST stores a :class:`SelfRef` node so the translator can emit
  Excel formulas that reference the previous cell.
- ``recurrence_sum`` — thin wrapper: ``x[t] = x[t-1] + increment[t]``.
- ``cumsum`` — running total; delegates to ``recurrence``.

These are the load-bearing time-series primitives; any alternative compute
renderer would need to lower ``SelfRef`` to ``WITH RECURSIVE`` / window
functions to support this shape of formula.
"""

from __future__ import annotations

import ast
import inspect
import math
import operator
import re
from datetime import date
from typing import Any, Callable, Dict, Optional, Tuple

from ..core.expr import Expr, Literal, SelfRef, VarRef
from ..core.variable import Variable


# ─── Safe formula evaluator ────────────────────────────────────────────
#
# Formula templates like ``"{prev} * (1 - {churn})"`` are user-authored
# expressions. We parse them with ``ast.parse(mode='eval')`` and then walk
# the resulting tree ourselves, allowing only a small set of operators,
# literals, and pre-registered function names. ``eval()`` is NOT used —
# this avoids every known Python-sandbox-escape pattern (attribute walks,
# dunder traversal, ``__import__`` tricks, list comprehensions, lambdas).
#
# Allowed: numeric literals, name references (from the caller-supplied
# bindings), arithmetic / comparison / unary operators, ternary
# ``x if c else y``, and calls to whitelisted functions below.
#
# Rejected: attribute access, subscripts, comprehensions, lambdas, walrus
# assignments, imports, arbitrary function calls.

_SAFE_BIN_OPS: Dict[type, Callable[[Any, Any], Any]] = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Mod: operator.mod,
    ast.Pow: operator.pow,
    ast.FloorDiv: operator.floordiv,
}

_SAFE_UNARY_OPS: Dict[type, Callable[[Any], Any]] = {
    ast.USub: operator.neg,
    ast.UAdd: operator.pos,
    ast.Not: operator.not_,
}

_SAFE_COMPARE_OPS: Dict[type, Callable[[Any, Any], bool]] = {
    ast.Eq: operator.eq, ast.NotEq: operator.ne,
    ast.Lt: operator.lt, ast.LtE: operator.le,
    ast.Gt: operator.gt, ast.GtE: operator.ge,
}

def _date_arg(x):
    """Coerce a recurrence binding to a ``date``.

    The previous period's value (``{prev}``) and any user-bound name in
    a date-bearing formula can arrive as ``date`` / ``datetime`` / ISO
    string. Centralised here so EDATE/EOMONTH/YEAR/MONTH/DAY share the
    same coercion semantics as the public Variable-level helpers in
    ``functions/dates.py``.
    """
    from .dates import _extract_date
    return _extract_date(x)


def _eomonth(d, months):
    from datetime import date
    from calendar import monthrange
    from dateutil.relativedelta import relativedelta
    shifted = _date_arg(d) + relativedelta(months=int(months))
    last_day = monthrange(shifted.year, shifted.month)[1]
    return date(shifted.year, shifted.month, last_day)


def _edate(d, months):
    from dateutil.relativedelta import relativedelta
    return _date_arg(d) + relativedelta(months=int(months))


_SAFE_FUNCTIONS: Dict[str, Any] = {
    'IF':    lambda c, a, b: a if c else b,
    'MAX':   max,
    'MIN':   min,
    'ABS':   abs,
    'ROUND': round,
    'INT':   math.floor,
    'MOD':   operator.mod,
    # Pure date arithmetic — safe to whitelist (no I/O, no attribute
    # access, deterministic on inputs). Mirror the public ``EDATE`` /
    # ``EOMONTH`` / ``YEAR`` / ``MONTH`` / ``DAY`` so a recurrence
    # template like ``"EDATE({prev}, 1)"`` evaluates correctly per
    # period and the rendered formula chains cell-to-cell in Excel.
    'EDATE':   _edate,
    'EOMONTH': _eomonth,
    'YEAR':    lambda d: _date_arg(d).year,
    'MONTH':   lambda d: _date_arg(d).month,
    'DAY':     lambda d: _date_arg(d).day,
}


def _safe_eval(node: ast.AST, names: Dict[str, Any]) -> Any:
    """Recursively interpret a parsed expression against a name bindings dict.

    Any AST node type not explicitly whitelisted raises ``ValueError``.
    """
    if isinstance(node, ast.Constant):
        if isinstance(node.value, (int, float, bool)):
            return node.value
        raise ValueError(
            f"Only numeric / boolean literals are allowed in formulas; "
            f"got {type(node.value).__name__}"
        )

    if isinstance(node, ast.Name):
        if node.id in names:
            return names[node.id]
        if node.id in _SAFE_FUNCTIONS:
            return _SAFE_FUNCTIONS[node.id]
        raise NameError(f"Unknown name in formula: {node.id!r}")

    if isinstance(node, ast.BinOp):
        fn = _SAFE_BIN_OPS.get(type(node.op))
        if fn is None:
            raise ValueError(f"Operator not allowed: {type(node.op).__name__}")
        return fn(_safe_eval(node.left, names), _safe_eval(node.right, names))

    if isinstance(node, ast.UnaryOp):
        # Distinct name from the BinOp branch's ``fn`` to keep type
        # narrowing clean — unary ops take one argument, binary takes two.
        unary_fn = _SAFE_UNARY_OPS.get(type(node.op))
        if unary_fn is None:
            raise ValueError(f"Unary operator not allowed: {type(node.op).__name__}")
        return unary_fn(_safe_eval(node.operand, names))

    if isinstance(node, ast.Compare):
        left = _safe_eval(node.left, names)
        for op, rhs in zip(node.ops, node.comparators):
            right = _safe_eval(rhs, names)
            cmp = _SAFE_COMPARE_OPS.get(type(op))
            if cmp is None:
                raise ValueError(f"Comparison not allowed: {type(op).__name__}")
            if not cmp(left, right):
                return False
            left = right
        return True

    if isinstance(node, ast.IfExp):
        cond = _safe_eval(node.test, names)
        return (_safe_eval(node.body, names) if cond
                else _safe_eval(node.orelse, names))

    if isinstance(node, ast.BoolOp):
        values = [_safe_eval(v, names) for v in node.values]
        if isinstance(node.op, ast.And):
            return all(values)
        if isinstance(node.op, ast.Or):
            return any(values)
        raise ValueError(f"Boolean operator not allowed: {type(node.op).__name__}")

    if isinstance(node, ast.Call):
        if not isinstance(node.func, ast.Name):
            raise ValueError("Only direct function calls allowed (no attribute access)")
        fn = (_SAFE_FUNCTIONS.get(node.func.id)
              or _SAFE_FUNCTIONS.get(node.func.id.upper()))
        if fn is None:
            raise NameError(f"Function not allowed in formulas: {node.func.id!r}")
        if node.keywords:
            raise ValueError("Keyword arguments not allowed in formula calls")
        return fn(*[_safe_eval(a, names) for a in node.args])

    raise ValueError(f"Expression not allowed in formulas: {type(node).__name__}")


# Converts ``"{name}"`` placeholders to bare identifiers so the template
# becomes parseable Python. ``"{prev} * (1 - {churn})"`` →
# ``"prev * (1 - churn)"``.
#: ``\w`` with re.UNICODE covers Cyrillic (and any other script) — the
#: product's models are written in Russian; an ASCII-only placeholder
#: silently left ``{п}`` unsubstituted and the parser then read the
#: braces as a Python set literal ("Expression not allowed: Set").
_PLACEHOLDER_RE = re.compile(r'\{([^\W\d]\w*)\}', re.UNICODE)


def _resolve_periods(periods, variables: Dict[str, Any]) -> Tuple[int, Optional[Variable]]:
    """Coerce ``periods`` to an int plus its source Variable (for AST).

    Resolution order:

    1. Explicit ``periods=N`` (int or Variable holding an int).
    2. The first list-valued input variable's length.

    Raises when neither is available — callers must be explicit about
    the period count to avoid ambiguity.
    """
    if periods is None:
        for var in variables.values():
            if isinstance(var, Variable) and var.var_type == 'list':
                return len(var._value), None
        raise ValueError(
            "recurrence() couldn't determine `periods`. Either:\n"
            "  - pass periods=N explicitly, e.g. recurrence(start, formula, periods=12), or\n"
            "  - pass a list-valued Variable as one of the template inputs."
        )

    if isinstance(periods, Variable):
        if not isinstance(periods._value, int):
            raise ValueError(
                f"recurrence() got a Variable for `periods` holding "
                f"{type(periods._value).__name__} {periods._value!r} — expected an int."
            )
        return periods._value, periods
    if not isinstance(periods, int):
        raise TypeError(
            f"recurrence() `periods` must be an int or Variable, "
            f"got {type(periods).__name__}."
        )
    return periods, None


def _per_period_values(variables: Dict[str, Any], periods: int) -> Dict[str, list]:
    """Align each variable to a per-period list of raw values.

    List-valued Variables pass through (Variables-of-Variables unwrapped).
    Scalars get broadcast to ``periods`` copies.
    """
    out: Dict[str, list] = {}
    for key, var in variables.items():
        if isinstance(var, Variable):
            if var.var_type == 'list' and isinstance(var._value, list):
                out[key] = [v._value if isinstance(v, Variable) else v for v in var._value]
            else:
                out[key] = [var._value] * periods
        else:
            out[key] = [var] * periods
    return out


def _eval_lambda(formula, start_value, periods: int) -> list:
    """Evaluate ``lambda prev, t: ...`` for periods 1..N-1; period 0 = start."""
    result = [start_value]
    prev = start_value
    for t in range(1, periods):
        try:
            current = formula(prev, t)
        except Exception as e:
            raise ValueError(f"Failed to evaluate lambda formula at period {t}: {e}")
        result.append(current)
        prev = current
    return result


def _eval_template(formula: str, start_value, periods: int,
                   var_values_per_period: Dict[str, list]) -> list:
    """Evaluate a string template against per-period values.

    Converts ``{name}`` placeholders into bare identifiers, parses the
    template once, and interprets the resulting AST per period against a
    names dict. No ``eval()`` — see :func:`_safe_eval` for the allowed
    expression surface.
    """
    parseable = _PLACEHOLDER_RE.sub(r'\1', formula)
    try:
        tree = ast.parse(parseable, mode='eval').body
    except SyntaxError as e:
        raise ValueError(f"Invalid formula syntax: {formula!r}: {e}") from e

    # New semantics (0.1.2): period 0 IS the start value verbatim. The
    # formula applies for periods 1..N-1 with prev = previous period.
    # Per-period variable values at index 0 are unused — kept aligned
    # with the result so callers can read `var[t]` and get the value
    # that was in effect at output period `t`.
    result = [start_value]
    prev = start_value
    for t in range(1, periods):
        names: Dict[str, Any] = {'prev': prev}
        for key, per_period in var_values_per_period.items():
            v = per_period[t]
            names[key] = v._value if isinstance(v, Variable) else v
        # Excel-error markers ABSORB before evaluation (string '+'
        # would silently concatenate them); ``None`` holes (§16.9 —
        # un-entered fact months) absorb through the TypeError below,
        # AFTER the structural validation inside ``_safe_eval`` has had
        # its say — a malformed template still teaches, but one missing
        # month must not kill the whole chain (the January case).
        if any(isinstance(v, str) and v.startswith('#')
               for v in names.values()):
            result.append('#VALUE!')
            prev = '#VALUE!'
            continue
        try:
            current = _safe_eval(tree, names)
        except TypeError as e:
            if any(v is None for v in names.values()):
                # A hole is absence, not failure — the chain goes
                # quiet (blank) from here rather than red (§16.9).
                result.append(None)
                prev = None
                continue
            raise ValueError(
                f"Error evaluating formula {formula!r} at period {t}: {e}"
            ) from e
        except (ValueError, NameError, ZeroDivisionError) as e:
            raise ValueError(
                f"Error evaluating formula {formula!r} at period {t}: {e}"
            ) from e
        result.append(current)
        prev = current
    return result


def _build_start_expr(start) -> Expr:
    """AST node for the recurrence start value."""
    if not isinstance(start, Variable):
        return Literal(start)
    # Prefer VarRef when the start has a referenceable identity; fall back
    # to its own AST for anonymous subscripts like ``sales_volume[0]``.
    if start.python_name or start._expr is None:
        return VarRef(start)
    return start._expr


def _build_selfref_expr(start, formula: str, variables: Dict[str, Any],
                        periods_var: Optional[Variable], periods: int) -> SelfRef:
    """Assemble a :class:`SelfRef` AST from the template + resolved variables."""
    template_variables: Dict[str, Expr] = {}
    for key, var in variables.items():
        template_variables[key] = VarRef(var) if isinstance(var, Variable) else Literal(var)
    periods_node = VarRef(periods_var) if periods_var is not None else periods
    return SelfRef(
        start=_build_start_expr(start),
        template=formula,
        variables=template_variables,
        periods=periods_node,
    )


def _source_label(var) -> str:
    """Friendly label for ``_source_code``: python_name, path leaf, or repr."""
    if isinstance(var, Variable):
        return getattr(var, 'python_name', None) or var.path.leaf
    return repr(var)


def _render_source_code(formula, start, variables: Dict[str, Any],
                        periods: Optional[int], periods_var: Optional[Variable],
                        is_lambda: bool) -> str:
    """Render a ``_source_code`` string close to the user's original call."""
    start_name = _source_label(start)
    if is_lambda:
        return f'recurrence({start_name}, <lambda>)'
    var_items = ', '.join(f"'{k}': {_source_label(v)}" for k, v in variables.items())
    periods_str = ''
    if periods is not None:
        name = _source_label(periods_var) if periods_var is not None else str(periods)
        periods_str = f', periods={name}'
    return f'recurrence({start_name}, "{formula}", {{{var_items}}}{periods_str})'


# ─── Public API ────────────────────────────────────────────────────


def recurrence(
    start: Any,
    formula: Optional[Any] = None,
    variables: Optional[Dict[str, Any]] = None,
    periods: Optional[Any] = None,
    display_name: Optional[str] = None,
    **kwargs: Any,
) -> Variable:
    """Build a list-valued Variable from a per-period recurrence.

    Pattern::

        x[0] = start
        x[t] = formula(prev=x[t-1], **vars[t])    for t = 1..N-1

    Args:
        start: Initial value (Variable or scalar).
        formula: String template (``"{prev} * (1 + {rate})"``) or a
            ``lambda prev, t: ...`` for Python-only evaluation.
        variables: Dict mapping template placeholders → Variables. Can also
            be passed as kwargs (``recurrence(start, fmt, rate=growth)``).
        periods: Explicit period count. Inferred from the first
            list-valued template Variable when omitted.
        display_name: Optional display label for the result.
        **kwargs: Template variables (``{name}`` substitutions). Excel
            styling goes through ``excel_props={...}`` instead of being
            forwarded here.

    Security:
        String templates are parsed with ``ast.parse`` and interpreted by
        a whitelisted AST walker — never ``eval()``'d. Allowed constructs:
        numeric literals, name references, arithmetic / comparison /
        unary / boolean operators, ternary ``x if c else y``, and calls to
        ``IF``, ``MAX``, ``MIN``, ``ABS``, ``ROUND``, ``INT``, ``MOD``.
        Anything else — attribute access, subscripts, comprehensions,
        lambdas, imports, arbitrary function calls — raises
        :class:`ValueError`. Lambda-form formulas run as plain Python and
        are only as safe as the callable the caller supplies.

    Lambdas have no Excel representation — the writer emits the computed
    values. String templates build a :class:`SelfRef` AST that translates
    to Excel formulas referencing the previous period's cell.
    """

    from ..core.shape import reject_axised
    from ..core.tracks import TrackValues as _Tracks
    reject_axised(start, "recurrence(start=)")
    for _k, _v in (variables or {}).items():
        reject_axised(_v, f"recurrence(variables[{_k!r}])")

    # Remaining kwargs are template variables (``{name}`` substitutions).
    if variables is None:
        variables = kwargs or {}

    # ─── track lift: N independent chains, one per coordinate ───
    # A roll-forward over a track-carrying driver rolls SEPARATELY in
    # each coordinate — the plan balance rolls from plan flows, the
    # actual balance from actual flows (§15.3.5; the live re-anchored
    # chain is the blend
    # law's third universe, §16.9 — P2, not this function).
    def _tracks_of(v: Any) -> Optional[_Tracks]:
        val = getattr(v, '_value', None) if isinstance(v, Variable) else None
        return val if isinstance(val, _Tracks) else None

    tracked = {k: t for k, t in
               ((k, _tracks_of(v)) for k, v in variables.items()) if t}
    start_tracks = _tracks_of(start)
    if tracked or start_tracks is not None:
        role_sets = [set(t.roles) for t in tracked.values()]
        if start_tracks is not None:
            role_sets.append(set(start_tracks.roles))
        if any(rs != role_sets[0] for rs in role_sets[1:]):
            raise ValueError(
                f"coordinate sets differ across recurrence operands: "
                f"{sorted(map(sorted, role_sets))} — slice or align "
                f"before the roll."
            )
        roles = (next(iter(tracked.values())).roles
                 if tracked else start_tracks.roles)  # type: ignore[union-attr]

        def _slice_operand(v: Any, role: str) -> Any:
            t = _tracks_of(v)
            if t is None:
                return v
            fiber = t[role]
            sliced = Variable()
            sliced._value = (list(fiber) if isinstance(fiber, list)
                             else fiber)
            sliced.var_type = ('list' if isinstance(fiber, list)
                               else 'scalar')
            return sliced

        # Period count: explicit wins; else the tracks' shared time
        # length (``_resolve_periods`` can't measure a Tracks value).
        periods_var: Optional[Variable] = None
        if periods is None:
            tl = next(
                (t.time_length for t in
                 ([*tracked.values()]
                  + ([start_tracks] if start_tracks else []))
                 if t.time_length is not None),
                None,
            )
            if tl is None:
                raise ValueError(
                    "recurrence() couldn't determine `periods` from "
                    "all-scalar tracks — pass periods=N explicitly."
                )
            periods_resolved = tl
        else:
            periods_resolved, periods_var = _resolve_periods(periods, {})

        per_role_values: Dict[str, list] = {}
        for role in roles:
            role_vars = {k: _slice_operand(v, role)
                         for k, v in variables.items()}
            role_start = _slice_operand(start, role)
            per_role_values[role] = recurrence(
                role_start, formula, role_vars, periods=periods_resolved,
            )._value

        is_lambda_ = callable(formula) and (
            inspect.isfunction(formula) or inspect.ismethod(formula)
        )
        if is_lambda_:
            result = Variable()
        else:
            result = Variable(formula=_build_selfref_expr(
                start, formula, variables, periods_var, periods_resolved,
            ))
        result._value = _Tracks(per_role_values)
        result.var_type = 'list'
        sample = next((f for f in per_role_values.values()
                       if isinstance(f, list)), [])
        result.value_type = ('float' if any(isinstance(v, float)
                                            for v in sample) else 'int')
        if is_lambda_:
            deps = [v for v in variables.values() if isinstance(v, Variable)]
            if isinstance(start, Variable):
                deps.append(start)
            result._dependency_refs = deps
        result._source_code = _render_source_code(
            formula, start, variables, periods_resolved, periods_var,
            is_lambda_,
        )
        if display_name:
            result._display_name = display_name
        return result

    is_lambda = callable(formula) and (inspect.isfunction(formula) or inspect.ismethod(formula))
    periods_resolved, periods_var = _resolve_periods(periods, variables)

    start_value = start._value if isinstance(start, Variable) else start

    # Evaluate in Python regardless of mode — user code expects ``._value`` populated.
    if is_lambda:
        values = _eval_lambda(formula, start_value, periods_resolved)
        recurrence_expr = None
    else:
        if formula is None:
            raise TypeError(
                "recurrence(start, formula, ...) requires a formula — pass a "
                "string template (e.g. '{prev} * (1 + {rate})') or a lambda."
            )
        values = _eval_template(
            formula, start_value, periods_resolved,
            _per_period_values(variables, periods_resolved),
        )
        recurrence_expr = _build_selfref_expr(
            start, formula, variables, periods_var, periods_resolved,
        )

    result = Variable(formula=recurrence_expr) if recurrence_expr is not None else Variable()

    result._value = values
    result.var_type = 'list'
    if any(isinstance(v, date) for v in values):
        result.value_type = 'datetime'  # e.g. recurrence(start, "EDATE(prev, 3)")
    elif any(isinstance(v, float) for v in values):
        result.value_type = 'float'
    else:
        result.value_type = 'int'

    # Lambda mode has no AST → dependency refs must be wired manually so the
    # dependency graph stays complete.
    if recurrence_expr is None:
        deps = []
        if isinstance(start, Variable):
            deps.append(start)
        if periods_var is not None:
            deps.append(periods_var)
        deps.extend(v for v in variables.values() if isinstance(v, Variable))
        result._dependency_refs = deps

    result._source_code = _render_source_code(
        formula, start, variables, periods_resolved, periods_var, is_lambda,
    )

    if display_name:
        result._display_name = display_name
    return result


def recurrence_sum(
    start: Any,
    increment: Any,
    periods: Optional[Any] = None,
    display_name: Optional[str] = None,
    **kwargs: Any,
) -> Variable:
    """Cumulative roll-forward — ``x[0] = start; x[t] = x[t-1] + increment[t]``.

    Thin wrapper around :func:`recurrence` with the fixed template
    ``"{prev} + {increment}"``. Equivalent to ``start + cumsum(increment)``
    but the recurrence pattern is explicit in the emitted formula.
    """
    return recurrence(
        start=start,
        formula="{prev} + {increment}",
        variables={'increment': increment},
        periods=periods,
        display_name=display_name,
        **kwargs,
    )


def cumsum(iterable: Variable) -> Variable:
    """Cumulative running total — ``[a, b, c] → [a, a+b, a+b+c]``.

    Delegates to :func:`recurrence` so the Excel output uses self-referencing
    formulas (each cell ``= prev + input``) instead of expanding the full
    sum range per cell.
    """
    from ..core.shape import reject_axised
    from ..core.tracks import TrackValues
    if isinstance(iterable, Variable) and isinstance(iterable._value, TrackValues):
        # Rank lifting: running total per coordinate.
        def _acc(track):
            if not isinstance(track, list):
                return track
            out, run = [], 0.0
            for v in track:
                # Holes and error markers stop the running total —
                # absence is quiet (None from here on), failure loud.
                if run is None or v is None:
                    run = None
                elif isinstance(v, str) or isinstance(run, str):
                    run = v if isinstance(v, str) else run
                else:
                    run += v
                out.append(run)
            return out
        result = Variable(formula=__import__(
            'modeleon.core.expr', fromlist=['MethodCall', 'VarRef']
        ).MethodCall(__import__(
            'modeleon.core.expr', fromlist=['VarRef']
        ).VarRef(iterable), 'cumsum', [], {}))
        result._value = TrackValues.lift(_acc, iterable._value)
        result.var_type = 'list'
        result._cumsum_source = iterable
        if getattr(iterable, '_unit', None) is not None:
            result._unit = iterable._unit  # a running total of ₸ is ₸
        return result
    reject_axised(iterable, "cumsum")
    if not isinstance(iterable, Variable):
        raise TypeError(
            f"cumsum() expects a Variable, got {type(iterable).__name__}. "
            f"For plain lists use numpy.cumsum or Python's itertools.accumulate. "
            f"Example: cumsum(mo.Variable([1, 2, 3]))."
        )

    if not isinstance(iterable._value, list):
        result = Variable(formula=VarRef(iterable))
        result._value = iterable._value
        result.var_type = 'scalar'
        result.value_type = iterable.value_type
        return result

    # New recurrence semantics: period 0 = start verbatim. Seed with
    # iterable[0] so the running total reads as [a, a+b, a+b+c, …].
    result = recurrence(
        start=iterable[0],
        formula="{prev} + {x}",
        variables={'x': iterable},
        periods=len(iterable._value),
    )

    # Source-code label: prefer the input's formula, then its python_name,
    # then its path leaf.
    if iterable.formula and not iterable.formula.startswith('_var_'):
        label = f"({iterable.formula})"
    else:
        label = iterable.python_name or iterable.path.leaf
    result._source_code = f"cumsum({label})"
    # Projection marker: a running total re-grains EXACTLY as the
    # cumsum of its re-grained (summed) input — cumsum(monthly)[q-end]
    # == cumsum(quarterly sums)[q]. Generic SelfRef recompute is
    # forbidden; this identity lets core.projection rebuild the row.
    result._cumsum_source = iterable
    if getattr(iterable, '_unit', None) is not None:
        result._unit = iterable._unit  # a running total of ₸ is ₸
    return result
