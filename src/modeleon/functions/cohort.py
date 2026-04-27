# SPDX-License-Identifier: Apache-2.0
"""Cohort-flavored finance helpers.

Currently: :func:`cohort_retention` — per-period active subscribers under
a flat retention rate, expressed as a self-referencing recurrence so the
Excel output is a live formula referencing the previous period's cell.
"""

from __future__ import annotations

from ..core.variable import Variable
from ._helpers import src
from .recurrence import recurrence


def cohort_retention(
    new_customers: Variable,
    retention_rate: "Variable | float",
) -> Variable:
    """Active subscribers under a flat retention rate.

    Expressed as the recurrence ``active[t] = active[t-1] * r + new[t]``,
    which is algebraically equivalent to summing every earlier cohort
    weighted by ``retention_rate`` raised to its age::

        Period 0: new[0]
        Period 1: new[1] + new[0] * r
        Period 2: new[2] + new[1] * r + new[0] * r**2

    Emitting the recurrence form (rather than a N-term sum per period)
    means the Excel output is ``=active_prev * r + new_t`` — one formula
    per cell, referencing the previous period.

    Args:
        new_customers: list-valued Variable — new acquisitions per period.
        retention_rate: scalar Variable — per-period retention (0..1).

    Returns:
        Variable with a list of active subscribers per period.
    """
    if not isinstance(new_customers, Variable):
        raise TypeError(
            f"cohort_retention() `new_customers` must be a Variable, "
            f"got {type(new_customers).__name__}. "
            f"Example: cohort_retention(mo.Variable([100, 150, 200]), rate)."
        )
    if not isinstance(retention_rate, Variable):
        raise TypeError(
            f"cohort_retention() `retention_rate` must be a Variable, "
            f"got {type(retention_rate).__name__}. "
            f"Example: cohort_retention(new_cust, mo.Variable(0.85))."
        )
    if not isinstance(new_customers._value, list):
        raise ValueError(
            "cohort_retention() `new_customers` must be a list Variable — "
            "new customers per period. Got a scalar."
        )
    if not isinstance(retention_rate._value, (int, float)):
        raise ValueError(
            "cohort_retention() `retention_rate` must be a scalar Variable "
            "(e.g., 0.85 for 85%). Got a list — use a single rate, not per-period."
        )

    # Period 0 = new[0] (no prior cohort to retain). Period t≥1 applies
    # the recurrence ``prev * r + new[t]``. Indexing the Variable yields
    # a sub-Variable backed by a Subscript AST so the Excel cell at
    # period 0 references new_customers[0] instead of inlining the
    # value as a literal.
    result = recurrence(
        start=new_customers[0],
        formula="{prev} * {r} + {n}",
        variables={'r': retention_rate, 'n': new_customers},
        periods=len(new_customers._value),
    )
    result._source_code = f"cohort_retention({src(new_customers)}, {src(retention_rate)})"
    return result
