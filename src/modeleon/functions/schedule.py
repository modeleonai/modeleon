# SPDX-License-Identifier: Apache-2.0
"""Piecewise-constant values keyed by DATE — ``mo.schedule``.

A "constant" like a tax rate lives in time and changes at dates the law
names. Spelling it as N hand-written cells or an ``IF`` chain lies about
its nature; a schedule states it:

    ндс = mo.schedule({'2025-01': 0.12, '2026-01': 0.16})

Dates, not positions — the spelling is grain-agnostic, survives a window
change, and reads like the regulation it encodes. The value materializes
when the Variable joins a windowed model (adoption resolves the ambient
``(start, grain, periods)``): each period carries the step in force at
its date. Aggregation defaults to ``mo.up('mean')`` — a rate re-grains
by (time-weighted) mean, never by sum.
"""

from __future__ import annotations

from typing import Any, Dict, Optional

from ..core.regrain import RegrainSpec, up
from ..core.variable import Variable


def schedule(
    steps: Dict[str, Any],
    display_name: Optional[str] = None,
    regrain: Optional[RegrainSpec] = None,
) -> Variable:
    """Build a piecewise-constant Variable from ``{date: value}`` steps.

    Args:
        steps: Mapping of ISO period labels (``'2026-01'`` for a monthly
            window, ``'2026'`` for yearly …) to the value in force FROM
            that period on. The window's first period must be covered:
            the earliest step must be at or before the window start.
        display_name: Optional display label.
        regrain: Aggregation rule; defaults to ``mo.up('mean')`` (a
            rate). Pass e.g. ``mo.up('sum')`` if the scheduled quantity
            is genuinely a flow.

    The step dates must land on period boundaries of the native grain —
    a mid-period change (``'2026-02-15'`` in a monthly model) is a
    teaching error rather than a silent shift.
    """
    if not isinstance(steps, dict) or not steps:
        raise TypeError(
            "mo.schedule({...}) needs a non-empty {date: value} dict, "
            "e.g. mo.schedule({'2025-01': 0.12, '2026-01': 0.16})."
        )
    for k in steps:
        if not isinstance(k, str) or not k.strip():
            raise TypeError(
                f"mo.schedule keys are ISO period labels (strings like "
                f"'2026-01'); got {k!r}."
            )

    var = Variable(
        display_name=display_name,
        regrain=regrain if regrain is not None else up('mean'),
    )
    # Ordered by label — ISO labels sort chronologically within a grain.
    var._schedule = dict(sorted(steps.items()))
    # A schedule holds its last value beyond its final step by nature.
    # ``extend=`` machinery lands with the track layer (P1); the intent
    # is stamped now so the alignment law can honor it without a
    # source migration.
    var._extend = 'hold'
    return var
