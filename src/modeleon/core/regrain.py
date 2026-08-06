# SPDX-License-Identifier: Apache-2.0
"""Re-grain — how a Variable moves between grains (explicit, per-variable).

A variable is defined at a native ``(start, grain)`` and is re-grainable: a
projection samples it at another grain. ``regrain=`` declares HOW, as inert
metadata on the Variable (this slice stores it; the projection driver
evaluates it later). The rule is always conceptually a callable
``fn(group) -> Variable``; named recipes, :func:`ratio`, and the (later)
``kind=`` sugar all reduce to that one form.

Constructors::

    regrain=mo.up('sum')                       # a flow
    regrain=mo.up('last')                       # a stock (period-end)
    regrain=mo.up(lambda g: mo.SUM(g.price * g.volume) / mo.SUM(g.volume))
    regrain=mo.up({('day', 'week'): 'last', ('week', 'month'): 'mean'})
    regrain=mo.ratio('rev', 'users')            # sum(num)/sum(den) per bucket
    regrain=mo.frozen()                          # grain-frozen (recurrence)

``down=`` (disaggregation) is reserved and raises; adding it later is purely
additive.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Optional, Tuple, Union

#: Built-in coarsening recipes. Each is sugar for a rule over the bucket's
#: own cells; the projection driver expands them. ``geometric`` is the
#: financially-correct rate re-grain (``PRODUCT(1+r)-1``), never a mean.
RECIPES = ('sum', 'first', 'last', 'mean', 'min', 'max', 'geometric')

Rule = Union[str, Callable[..., Any]]
Transition = Tuple[str, str]


@dataclass(frozen=True)
class Ratio:
    """A ratio-of-flows re-grain: ``sum(num)/sum(den)`` per bucket (VWAP /
    margin / ARPU). ``num`` and ``den`` name sibling sources."""

    num: str
    den: str


@dataclass(frozen=True)
class RegrainSpec:
    """Inert re-grain metadata stored on ``Variable._regrain``.

    ``default`` is the rule for any coarsening; ``overrides`` keys it per
    ``(from_grain, to_grain)`` transition; ``frozen`` opts the series out of
    projection re-grain. No aggregation math lives here — that is the
    projection driver's job (a later slice).
    """

    default: Optional[Union[Rule, Ratio]] = None
    overrides: Dict[Transition, Rule] = field(default_factory=dict)
    frozen: bool = False


def _normalize_rule(rule: Any) -> Union[Rule, Ratio]:
    if callable(rule) or isinstance(rule, Ratio):
        return rule
    if isinstance(rule, str):
        if rule not in RECIPES:
            raise ValueError(
                f"Unknown re-grain recipe {rule!r}; known recipes are "
                f"{', '.join(RECIPES)}, or pass a callable fn(group) -> Variable."
            )
        return rule
    raise TypeError(
        f"A re-grain rule must be a recipe name or a callable; got "
        f"{type(rule).__name__}."
    )


def _normalize_overrides(rules: Dict[Any, Any]) -> Dict[Transition, Rule]:
    out: Dict[Transition, Rule] = {}
    for key, rule in rules.items():
        if not (isinstance(key, tuple) and len(key) == 2):
            raise ValueError(
                f"Re-grain override keys must be (from_grain, to_grain) tuples; "
                f"got {key!r}."
            )
        out[key] = _normalize_rule(rule)  # type: ignore[assignment]
    return out


def up(rule: Union[Rule, Dict[Any, Any], None] = None, *,
       down: Any = None) -> RegrainSpec:
    """Build a coarsening (roll-up) re-grain spec.

    ``rule`` is a recipe name, a callable ``fn(group) -> Variable``, or a
    ``{(from_grain, to_grain): rule}`` dict for per-transition rules. ``down=``
    (disaggregation) is reserved and raises until that feature ships.
    """
    if down is not None:
        raise NotImplementedError(
            "down= (disaggregation, coarse->fine) is not implemented yet; only "
            "up/coarsen is supported. Re-grain a finer-grained source instead."
        )
    if rule is None:
        raise ValueError(
            "mo.up() needs a rule: a recipe name, a callable fn(group) -> "
            "Variable, or a {(from, to): rule} dict."
        )
    if isinstance(rule, dict):
        return RegrainSpec(overrides=_normalize_overrides(rule))
    return RegrainSpec(default=_normalize_rule(rule))


def frozen(values: Optional[str] = None) -> RegrainSpec:
    """Mark a series grain-frozen: its FORMULA only means anything at its
    native grain (recurrence / roll-forward, ``x[t]=f(x[t-1])``).

    Frozen splits formula from values. ``mo.frozen(values='last')`` says the
    VALUE series still has coarse meaning — quarter-end cash is the last
    monthly close — so a projection re-grains the computed values by that
    rule while never re-evaluating the recurrence at the target grain. Bare
    ``mo.frozen()`` is the absolute form for series with no grain-free
    meaning at all (a period counter, the date spine): projecting one raises.
    """
    if values is None:
        return RegrainSpec(frozen=True)
    if not isinstance(values, str) or values not in RECIPES:
        raise ValueError(
            f"mo.frozen(values=...) takes a named recipe "
            f"({', '.join(RECIPES)}); got {values!r}."
        )
    return RegrainSpec(frozen=True, default=values)


def ratio(num: str, den: str) -> RegrainSpec:
    """A ratio-of-flows re-grain: ``sum(num)/sum(den)`` per target bucket.

    The right form for VWAP / blended margin / ARPU — a ratio must re-grain as
    sum-over-sum, never as a mean of the per-period ratios.
    """
    return RegrainSpec(default=Ratio(num, den))


def apply_recipe(bucket: list, recipe: str,
                 weights: Optional[list] = None) -> Any:
    """Reduce one bucket of native cells to a single coarse value via a named
    recipe. (Callable and :class:`Ratio` rules are evaluated by the projection
    driver, not here.)

    ``weights`` are the periods' day-counts (from
    :func:`modeleon.core.time.period_weights`) — periods are not equal (a
    quarter mixes 28- and 31-day months; a stub period may be days long), so
    ``mean`` is **time-weighted** when weights are given; ``None`` means
    uniform weights. Order-insensitive recipes ignore weights: a flow sums
    regardless of period length, and ``geometric`` compounds per period (each
    rate already applies to its own period, whatever its length).

    Error cells absorb: a bucket containing an Excel error token (a ``'#'``-
    prefixed string, the same convention the operator chokepoint uses) reduces
    to that token. Domain failures produce tokens, never raise: ``geometric``
    with any ``1 + x <= 0`` is ``'#NUM!'``; a weighted ``mean`` over zero
    total weight is ``'#DIV/0!'`` — one bad bucket is one error cell, not a
    projection-wide crash.
    """
    if not bucket:
        raise ValueError("Cannot re-grain an empty bucket.")
    if weights is not None and len(weights) != len(bucket):
        raise ValueError(
            f"Re-grain weights length {len(weights)} does not match bucket "
            f"length {len(bucket)}."
        )
    for value in bucket:
        if isinstance(value, str) and value.startswith('#'):
            return value                       # absorbing error token
    if recipe == 'sum':
        return sum(bucket)
    if recipe == 'first':
        return bucket[0]
    if recipe == 'last':
        return bucket[-1]
    if recipe == 'mean':
        if weights is None:
            return sum(bucket) / len(bucket)
        total = sum(weights)
        if total == 0:
            return '#DIV/0!'
        return sum(v * w for v, w in zip(bucket, weights)) / total
    if recipe == 'min':
        return min(bucket)
    if recipe == 'max':
        return max(bucket)
    if recipe == 'geometric':
        product = 1.0
        for value in bucket:
            if 1 + value <= 0:
                return '#NUM!'                 # -100% or worse: no real compound
            product *= (1 + value)
        return product - 1
    raise ValueError(
        f"Re-grain recipe {recipe!r} is not a numeric reducer; known reducers "
        f"are {', '.join(RECIPES)}."
    )
