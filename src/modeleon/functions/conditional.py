# SPDX-License-Identifier: Apache-2.0
"""Conditional helpers — ``IF`` (Variable-aware, renders ``=IF(cond, a, b)``)."""

from __future__ import annotations

from typing import Any

from ..core.variable import Variable
from ._helpers import make_func_var, text_operand


def IF(condition: Any, value_if_true: Any, value_if_false: Any) -> Variable:
    """Variable-aware ternary that renders as Excel's ``=IF(cond, a, b)``.

    Accepts Variables or plain values for each argument. Condition may be a
    scalar boolean / Variable, or a list Variable for element-wise IF.

    Example::

        bonus = mo.IF(revenue >= target, 50_000, 0)
        rating = mo.IF(ebitda > 200_000, "Excellent",
                       mo.IF(ebitda > 100_000, "Good", "Poor"))
    """
    cond_value, cond_expr = text_operand(condition)
    true_value, true_expr = text_operand(value_if_true)
    false_value, false_expr = text_operand(value_if_false)

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
