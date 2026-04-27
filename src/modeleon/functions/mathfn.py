# SPDX-License-Identifier: Apache-2.0
"""Element-wise math helpers — ``ABS``, ``ROUND``, ``INT``, ``MOD``.

Named ``mathfn`` (not ``math``) to avoid shadowing Python's stdlib ``math``.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula. ``mo.ABS(-5)`` produces ``=ABS(-5)`` in Excel, not a bare ``5`` —
the formula is the product.
"""

from __future__ import annotations

import builtins
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


def ROUND(value: Numeric, digits: Numeric = 0) -> Variable:
    """Round a Variable or scalar to ``digits`` decimal places (renders ``=ROUND(...)``)."""
    raw, expr = operand(value)
    raw_digits, digits_expr = operand(digits)
    d = int(raw_digits)
    calc, var_type = _apply(lambda v: builtins.round(v, d), raw)
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
