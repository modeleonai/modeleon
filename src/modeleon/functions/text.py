# SPDX-License-Identifier: Apache-2.0
"""Text helpers — ``LEN``, ``UPPER``, ``LOWER``, ``CONCAT``.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula — even when all inputs are plain strings. ``mo.UPPER("hello")``
produces ``=UPPER("hello")`` in Excel, not a bare ``"HELLO"``.
"""

from __future__ import annotations

from typing import Any, Callable, Union

from ..core.variable import Variable
from ._helpers import make_func_var, src, text_operand

TextOperand = Union[Variable, str]


def _unary_text_func(
    func_name: str,
    py_fn: Callable[[Any], Any],
    result_value_type: str,
) -> Callable[[TextOperand], Variable]:
    """Build a Variable-aware string helper like LEN / UPPER / LOWER.

    Always returns a Variable, regardless of whether input is a Variable or
    a plain string.
    """
    def helper(value: TextOperand) -> Variable:
        raw, expr = text_operand(value)
        if isinstance(raw, list):
            calc: Any = [py_fn(v) if v is not None else None for v in raw]
            var_type = 'list'
        else:
            calc = py_fn(raw) if raw is not None else None
            var_type = 'scalar'
        return make_func_var(
            func_name, [expr], calc, result_value_type, var_type,
            source_code=f"{func_name}({src(value)})",
        )

    return helper


LEN = _unary_text_func("LEN", lambda s: len(str(s)), 'int')
LEN.__doc__ = "Character length of a string Variable or plain string (renders ``=LEN(...)``)."

UPPER = _unary_text_func("UPPER", lambda s: str(s).upper(), 'string')
UPPER.__doc__ = "Uppercase a string Variable or plain string (renders ``=UPPER(...)``)."

LOWER = _unary_text_func("LOWER", lambda s: str(s).lower(), 'string')
LOWER.__doc__ = "Lowercase a string Variable or plain string (renders ``=LOWER(...)``)."


def _concat_piece(v: Any) -> str:
    """Render one operand value for concatenation, Excel-style.

    ``None`` → empty; an integer-valued float (``2022.0``) → ``"2022"``
    so period labels read ``"1Q 2022"`` not ``"1Q 2022.0"``.
    """
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def CONCAT(*args: TextOperand) -> Variable:
    """Concatenate Variables and/or plain strings (renders ``=CONCAT(...)``).

    Element-wise over list (period-indexed) operands, so
    ``mo.CONCAT(quarter, "Q ", year, " ", phase)`` over per-period
    Variables yields per-period labels (``"1Q 2022 Ф"``, …). Always
    returns a Variable carrying the ``=CONCAT(...)`` formula.
    """
    if not args:
        raise ValueError(
            "CONCAT() needs at least one argument. "
            "Syntax: mo.CONCAT(prefix, name)  or  mo.CONCAT(\"Q1 \", \"2025\"). "
            "Got 0 arguments."
        )

    raws, expr_args = [], []
    for a in args:
        raw, expr = text_operand(a)
        raws.append(raw)
        expr_args.append(expr)

    if any(isinstance(r, list) for r in raws):
        n = max(len(r) for r in raws if isinstance(r, list))

        def at(r: Any, i: int) -> Any:
            if isinstance(r, list):
                return r[i] if i < len(r) else r[-1]
            return r

        value: Any = ["".join(_concat_piece(at(r, i)) for r in raws) for i in range(n)]
        var_type = 'list'
    else:
        value = "".join(_concat_piece(r) for r in raws)
        var_type = 'scalar'

    source_code = "CONCAT(" + ", ".join(src(a) for a in args) + ")"
    return make_func_var(
        "CONCAT", expr_args, value, 'string', var_type,
        source_code=source_code,
    )
