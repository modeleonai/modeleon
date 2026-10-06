# SPDX-License-Identifier: Apache-2.0
"""Re-grain — how a Variable moves between grains (explicit, per-variable).

A variable is defined at a native ``(start, grain)`` and is re-grainable: a
projection samples it at another grain. ``regrain=`` declares HOW, as inert
metadata on the Variable (the Variable only stores it; the projection driver
evaluates it). The rule is always conceptually a callable
``fn(group) -> Variable``; named recipes and :func:`ratio` both reduce to
that one form.

Constructors::

    regrain=mo.up('sum')                       # a flow
    regrain=mo.up('last')                       # a stock (period-end)
    regrain=mo.up(lambda g: mo.SUM(g.price * g.volume) / mo.SUM(g.volume))
    regrain=mo.up({('day', 'week'): 'last', ('week', 'month'): 'mean'})
    regrain=mo.ratio('rev', 'users')            # sum(num)/sum(den) per bucket
    regrain=mo.frozen()                          # grain-frozen (recurrence)

``down=`` (disaggregation) is reserved and raises; only coarsening is
supported.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date as _date
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
    projection driver's job.
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
    (disaggregation) is reserved and raises — it is not supported.
    """
    if down is not None:
        raise NotImplementedError(
            "down= (disaggregation, coarse->fine) is not supported; only "
            "up/coarsen is. Re-grain a finer-grained source instead."
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


def hole_value(series: Any) -> Any:
    """What a hole counts as when ``series`` is folded: ``0``, as a blank
    cell in a formula - or ``None`` in a line of dates or words, where a
    blank has no zero. A line with tracks is judged on all of them: its
    actuals not entered yet are no less a line of dates."""
    if hasattr(series, 'roles'):
        values = [v for role in series.roles
                  for v in (series[role] if isinstance(series[role], list) else [series[role]])]
    else:
        values = series if isinstance(series, list) else [series]
    return None if any(
        isinstance(value, _date) or (isinstance(value, str) and not value.startswith('#'))
        for value in values) else 0


#: ``blank`` not given: the values at hand decide what a hole counts as.
LINE_UNKNOWN: Any = object()


def apply_recipe(bucket: list, recipe: str,
                 weights: Optional[list] = None, *, blank: Any = LINE_UNKNOWN) -> Any:
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

    A hole (``None``, an un-entered period) counts as zero, as a blank cell
    does in the book's formula: a quarter with one month entered sums that
    month, and ``geometric`` compounds the blank month at 0%. ``mean``,
    ``min`` and ``max`` skip holes, as Excel's ``AVERAGE`` / ``MIN`` /
    ``MAX`` skip blank cells (a mean over the entered periods' days), and a
    bucket of holes alone is ``0``. Positional recipes (``first`` /
    ``last``) read only their own period, matching the cell reference they
    emit: a hole there is ``0``, a sibling's error is not theirs. A line of
    dates or words has no zero: a bucket with a hole stays ``None`` there
    (a positional pick of an entered period excepted). ``blank`` is what a
    hole counts as (:func:`hole_value` of the whole line); left out, the
    bucket alone decides.

    Error cells absorb into aggregating recipes: a bucket containing an
    Excel error token (a ``'#'``-prefixed string, the same convention the
    operator chokepoint uses) reduces to that token. Domain failures produce
    tokens, never raise: ``geometric`` with any ``1 + x <= 0`` is
    ``'#NUM!'``; a weighted ``mean`` over zero total weight is ``'#DIV/0!'``
    - one bad bucket is one error cell, not a projection-wide crash.
    """
    if not bucket:
        raise ValueError("Cannot re-grain an empty bucket.")
    if weights is not None and len(weights) != len(bucket):
        raise ValueError(
            f"Re-grain weights length {len(weights)} does not match bucket "
            f"length {len(bucket)}."
        )
    if blank is LINE_UNKNOWN:
        blank = hole_value(bucket)
    if recipe in ('first', 'last'):
        picked = bucket[0] if recipe == 'first' else bucket[-1]
        return blank if picked is None else picked
    for value in bucket:
        if isinstance(value, str) and value.startswith('#'):
            return value                       # absorbing error token
    if blank is None and any(value is None for value in bucket):
        return None
    if recipe in ('min', 'max'):
        entered = [value for value in bucket if value is not None]
        if not entered:
            return blank
        return min(entered) if recipe == 'min' else max(entered)
    if recipe == 'mean':
        # Over the periods entered, as Excel's AVERAGE skips a blank cell.
        pairs = [(value, 1 if weights is None else weights[i])
                 for i, value in enumerate(bucket) if value is not None]
        if not pairs:
            return blank
        total = sum(w for _v, w in pairs)
        if total == 0:
            return '#DIV/0!'
        return sum(v * w for v, w in pairs) / total
    values = [0 if value is None else value for value in bucket]
    if recipe == 'sum':
        return sum(values)
    if recipe == 'geometric':
        product = 1.0
        for value in values:
            if 1 + value <= 0:
                return '#NUM!'                 # -100% or worse: no real compound
            product *= (1 + value)
        return product - 1
    raise ValueError(
        f"Re-grain recipe {recipe!r} is not a numeric reducer; known reducers "
        f"are {', '.join(RECIPES)}."
    )
