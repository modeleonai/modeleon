# SPDX-License-Identifier: Apache-2.0
"""Conditional / logical helpers — ``IF``, ``AND``, ``OR``, ``NOT``, ``CHOOSE``.

All are Variable-aware and element-wise over list (period-indexed)
operands, and render as their Excel equivalents (``=IF(...)``,
``=AND(...)``, ``=CHOOSE(...)``, …).
"""

from __future__ import annotations

from typing import Any

from ..core.variable import Variable
from ._helpers import make_func_var, operand, text_operand


def _value_type_of(sample: Any) -> str:
    """Modeleon value-type tag for a computed scalar sample."""
    if isinstance(sample, bool):
        return 'bool'
    if isinstance(sample, int):
        return 'int'
    if isinstance(sample, float):
        return 'float'
    if isinstance(sample, str):
        return 'string'
    return 'float'


def IF(condition: Any, value_if_true: Any, value_if_false: Any) -> Variable:
    """Variable-aware ternary that renders as Excel's ``=IF(cond, a, b)``.

    Accepts Variables or plain values for each argument. Condition may be a
    scalar boolean / Variable, or a list Variable for element-wise IF.

    Example::

        bonus = mo.IF(revenue >= target, 50_000, 0)
        rating = mo.IF(ebitda > 200_000, "Excellent",
                       mo.IF(ebitda > 100_000, "Good", "Poor"))
    """
    from ..core.shape import reject_axised
    for _arg, _where in ((condition, "IF(condition)"),
                         (value_if_true, "IF(value_if_true)"),
                         (value_if_false, "IF(value_if_false)")):
        reject_axised(_arg, _where)
    cond_value, cond_expr = text_operand(condition)
    true_value, true_expr = text_operand(value_if_true)
    false_value, false_expr = text_operand(value_if_false)

    from ..core.tracks import TrackValues
    if any(isinstance(v, TrackValues) for v in (cond_value, true_value, false_value)):
        # Rank lifting: the ternary evaluates per coordinate; plain
        # operands broadcast into every role (mismatch law inside lift).
        def _tern(c, t, f):
            if isinstance(c, list):
                return [
                    (t[i] if isinstance(t, list) else t) if cv
                    else (f[i] if isinstance(f, list) else f)
                    for i, cv in enumerate(c)
                ]
            return t if c else f

        lifted = TrackValues.lift(_tern, cond_value, true_value, false_value)
        result = make_func_var(
            "IF", [cond_expr, true_expr, false_expr], lifted, 'float',
            'list' if lifted.time_length else 'scalar',
        )
        return result

    if isinstance(cond_value, list):
        value = [
            (true_value[i] if isinstance(true_value, list) else true_value) if c
            else (false_value[i] if isinstance(false_value, list) else false_value)
            for i, c in enumerate(cond_value)
        ]
        var_type = 'list'
    else:
        value = true_value if cond_value else false_value
        var_type = 'scalar'

    sample = value[0] if isinstance(value, list) and value else value
    if isinstance(sample, bool):
        value_type = 'bool'
    elif isinstance(sample, int):
        value_type = 'int'
    elif isinstance(sample, float):
        value_type = 'float'
    elif isinstance(sample, str):
        value_type = 'string'
    else:
        value_type = 'float'

    result = make_func_var(
        "IF", [cond_expr, true_expr, false_expr], value, value_type, var_type,
    )
    result._source_code = result.formula
    return result


def _logical(func_name: str, reduce_fn, *args: Any) -> Variable:
    """Shared scaffold for ``AND`` / ``OR`` — element-wise over list operands.

    Each argument is a boolean-ish Variable (a comparison result) or a
    plain value. When any operand is a list, the result is element-wise
    (per period); otherwise scalar. ``reduce_fn`` is :func:`all` or
    :func:`any`.
    """
    raws, exprs = [], []
    for a in args:
        v, e = text_operand(a)
        raws.append(v)
        exprs.append(e)

    if any(isinstance(r, list) for r in raws):
        n = max(len(r) for r in raws if isinstance(r, list))

        def at(r: Any, i: int) -> Any:
            if isinstance(r, list):
                return r[i] if i < len(r) else r[-1]
            return r

        value: Any = [reduce_fn(bool(at(r, i)) for r in raws) for i in range(n)]
        var_type = 'list'
    else:
        value = reduce_fn(bool(r) for r in raws)
        var_type = 'scalar'

    result = make_func_var(func_name, exprs, value, 'bool', var_type)
    result._source_code = result.formula
    return result


def AND(*args: Any) -> Variable:
    """Logical AND (renders ``=AND(...)``). Element-wise over list operands.

    ``mo.AND(start >= invest_start, balance < 1)`` → per-period boolean
    Variable, true where both conditions hold.
    """
    return _logical("AND", all, *args)


def OR(*args: Any) -> Variable:
    """Logical OR (renders ``=OR(...)``). Element-wise over list operands."""
    return _logical("OR", any, *args)


def NOT(value: Any) -> Variable:
    """Logical NOT (renders ``=NOT(...)``). Element-wise over a list operand."""
    raw, expr = text_operand(value)
    if isinstance(raw, list):
        calc: Any = [not bool(v) for v in raw]
        var_type = 'list'
    else:
        calc = not bool(raw)
        var_type = 'scalar'
    result = make_func_var("NOT", [expr], calc, 'bool', var_type)
    result._source_code = result.formula
    return result


def ISBLANK(value: Any) -> Variable:
    """TRUE when the operand is empty (renders ``=ISBLANK(...)``).

    The Excel-native way to ask "was this ever filled in?", and the only
    way to write an OPTIONAL date without lying about it. A leaving date
    that has not happened yet has to read as an empty cell — filling it
    with the end of the window says "dismissed on 31 December", which an
    accountant will act on. Element-wise over a list operand.
    """
    raw, expr = text_operand(value)
    if isinstance(raw, list):
        calc: Any = [v is None or v == '' for v in raw]
        var_type = 'list'
    else:
        calc = raw is None or raw == ''
        var_type = 'scalar'
    result = make_func_var("ISBLANK", [expr], calc, 'bool', var_type)
    result._source_code = result.formula
    return result


def CHOOSE(index: Any, *choices: Any) -> Variable:
    """Pick the ``index``-th value, 1-based (renders ``=CHOOSE(...)``).

    The classic scenario/phase switch — ``mo.CHOOSE(scenario, base, bull)``
    or, element-wise, ``mo.CHOOSE(phase_flag + 1, "Ф", "П")`` where
    ``phase_flag`` is a per-period Variable. Out-of-range indices yield
    ``#VALUE!`` (matching Excel).
    """
    idx_raw, idx_expr = operand(index)
    pairs = [text_operand(c) for c in choices]
    choice_raws = [p[0] for p in pairs]
    choice_exprs = [p[1] for p in pairs]

    def pick(i_val: Any, period: int) -> Any:
        try:
            k = int(i_val) - 1
        except (TypeError, ValueError):
            return '#VALUE!'
        if not 0 <= k < len(choice_raws):
            return '#VALUE!'
        r = choice_raws[k]
        if isinstance(r, list):
            return r[period] if period < len(r) else r[-1]
        return r

    if isinstance(idx_raw, list):
        value: Any = [pick(iv, p) for p, iv in enumerate(idx_raw)]
        var_type = 'list'
    else:
        value = pick(idx_raw, 0)
        var_type = 'scalar'

    sample = value[0] if isinstance(value, list) and value else value
    result = make_func_var(
        "CHOOSE", [idx_expr, *choice_exprs], value, _value_type_of(sample), var_type,
    )
    result._source_code = result.formula
    return result
