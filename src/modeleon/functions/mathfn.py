# SPDX-License-Identifier: Apache-2.0
"""Element-wise math helpers — ``ABS``, ``ROUND``, ``INT``, ``MOD``.

Named ``mathfn`` (not ``math``) to avoid shadowing Python's stdlib ``math``.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula. ``mo.ABS(-5)`` produces ``=ABS(-5)`` in Excel, not a bare ``5`` —
the formula is the product.
"""

from __future__ import annotations

import builtins
import decimal
import math
from typing import Any, Union

from ..core.variable import Variable
from ._helpers import make_func_var, operand, src

Number = Union[int, float]
Numeric = Union[Variable, Number]


def _apply(py_fn, raw) -> tuple[Any, str]:
    """Apply ``py_fn`` element-wise over list values, passthrough None."""
    if isinstance(raw, list):
        return [py_fn(v) if v is not None else None for v in raw], 'list'
    return (py_fn(raw) if raw is not None else None), 'scalar'


def ABS(value: Numeric) -> Variable:
    """Absolute value (renders ``=ABS(...)``). Scalar or element-wise on lists.

    Accepts a Variable or a plain number. Always returns a Variable.
    """
    raw, expr = operand(value)
    calc, var_type = _apply(builtins.abs, raw)
    value_type = value.value_type if isinstance(value, Variable) else 'float'
    return make_func_var(
        "ABS", [expr], calc, value_type, var_type,
        source_code=f"ABS({src(value)})",
    )


#: No limit on digits or exponent, so a very long whole number or an
#: absurd ``digits`` (``ROUND(x, -500)`` is 0) never raises a decimal error.
_ROUND_CONTEXT = decimal.Context(
    prec=decimal.MAX_PREC, rounding=decimal.ROUND_HALF_UP,
    Emax=decimal.MAX_EMAX, Emin=decimal.MIN_EMIN,
)
#: The 15 significant digits a cell shows, a half in the 16th going
#: away from zero.
_SHOWN_CONTEXT = decimal.Context(prec=15, rounding=decimal.ROUND_HALF_UP)


def _spreadsheet_round(v, digits: int):
    """Round ``v`` the way the spreadsheet's ``ROUND`` does.

    Two differences from Python's ``round``: a half goes AWAY FROM ZERO
    (2.5 -> 3, -2.5 -> -3; Python sends it to the even digit, 2), and
    the number rounded is the one a cell shows — the float to 15
    significant digits — not its binary value. ``1.005`` is held as
    1.00499999…, so Python rounds it to 1.0 while the sheet shows 1.005
    and rounds it to 1.01. (``ROUND_HALF_UP`` in :mod:`decimal` is
    away from zero.) A number of 16 or more whole digits has no decimals
    left to show and is rounded as held, so its whole digits never move.
    Anything that is not a finite number goes to Python's ``round`` as
    before.
    """
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return builtins.round(v, digits)
    if isinstance(v, int):
        if digits >= 0:
            return v                            # a whole number has no decimals
        shown = decimal.Decimal(v)
    elif not math.isfinite(v):
        return builtins.round(v, digits)
    elif abs(v) >= 1e15:
        shown = decimal.Decimal(v)              # exactly as held
    else:
        shown = _SHOWN_CONTEXT.plus(decimal.Decimal(v))
    if digits >= -shown.as_tuple().exponent:
        rounded = shown                         # no digit past the cut-off
    else:
        step = decimal.Decimal(1).scaleb(-digits, context=_ROUND_CONTEXT)
        rounded = shown.quantize(step, context=_ROUND_CONTEXT)
    return int(rounded) if isinstance(v, int) else float(rounded)


def ROUND(value: Numeric, digits: Numeric = 0) -> Variable:
    """Round a Variable or scalar to ``digits`` decimal places (renders ``=ROUND(...)``).

    Computes what the spreadsheet computes: halves round away from zero
    (``ROUND(2.5, 0) == 3``), on the number as a cell shows it
    (``ROUND(1.005, 2) == 1.01``). Negative ``digits`` round to tens,
    hundreds, … (``ROUND(1250, -2) == 1300``).
    """
    raw, expr = operand(value)
    raw_digits, digits_expr = operand(digits)
    d = int(raw_digits)
    calc, var_type = _apply(lambda v: _spreadsheet_round(v, d), raw)
    value_type = value.value_type if isinstance(value, Variable) else 'float'
    return make_func_var(
        "ROUND", [expr, digits_expr], calc, value_type, var_type,
        source_code=f"ROUND({src(value)}, {d})",
    )


def INT(value: Numeric) -> Variable:
    """Floor a Variable or scalar to the nearest integer (renders ``=INT(...)``).

    Matches Excel's ``INT`` (floor toward -∞), not Python's ``int()``
    (truncate toward 0). So ``INT(-2.5) == -3``.
    """
    raw, expr = operand(value)
    calc, var_type = _apply(math.floor, raw)
    return make_func_var(
        "INT", [expr], calc, 'int', var_type,
        source_code=f"INT({src(value)})",
    )


def MOD(number: Numeric, divisor: Numeric) -> Variable:
    """Remainder after division (renders ``=MOD(number, divisor)``).

    For the usual case, the ``%`` operator already compiles to ``MOD(...)``.
    Use this when you want the explicit function call form in the Excel output.
    """
    num_val, num_expr = operand(number)
    div_val, div_expr = operand(divisor)
    calc, var_type = _apply(lambda v: v % div_val, num_val)
    return make_func_var(
        "MOD", [num_expr, div_expr], calc, 'float', var_type,
        source_code=f"MOD({src(number)}, {src(divisor)})",
    )
