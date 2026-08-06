# SPDX-License-Identifier: Apache-2.0
"""Out-of-extent behavior — ``extend=`` (§16.5).

A partial series is legal when its author DECLARES what lies beyond the
values given. One word resolves the trilemma a bare prefix cannot:

    капекс = mo.Variable([500, 500, 500], extend=mo.zero())   # a flow
    ставка = mo.Variable([0.10, 0.10, 0.12], extend=mo.hold())  # a rate

``zero`` — the quantity is absent beyond its values (flows, one-time
items). ``hold`` — the last value stays in force (rates, levels).
``none`` — reading beyond the values is an error; reserved for the
track layer's absent/NA semantics (declaring it today parses but does
not yet materialize a partial series).

v1 materializes the declared extension at ADOPTION, out to the ambient
window — the same moment a schedule materializes and the extent law
runs. Date-aligned arithmetic between unequal extents (§16.5 proper)
lands with the track layer; materialization makes the declared intent
computable today without it.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ExtendRule:
    """The declared out-of-extent behavior of a partial series."""

    kind: str  # 'zero' | 'hold' | 'none'

    def __repr__(self) -> str:  # renders like its constructor
        return f"mo.{self.kind}()"


def zero() -> ExtendRule:
    """Beyond its values the quantity is ABSENT — continue with 0.0
    (flows: costs that ended, one-time items)."""
    return ExtendRule('zero')


def hold() -> ExtendRule:
    """Beyond its values the LAST value stays in force (rates,
    levels, prices)."""
    return ExtendRule('hold')


def none() -> ExtendRule:
    """Reading beyond the values is an error — no honest continuation
    exists. Reserved for the track layer's absent/NA semantics."""
    return ExtendRule('none')
