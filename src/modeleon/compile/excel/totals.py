# SPDX-License-Identifier: Apache-2.0
"""Subtotal columns interleaved with the native periods.

The book form every finance team builds by hand:

    Jan   Feb   Mar   (Q1)   Apr … Dec   (Q4)   (year)

``ExcelView(timeline={'totals': ['quarter', 'year']})`` declares it;
this module owns the geometry (which physical column each month and
each bucket lands in) and the CONTENT rule for a bucket cell:

* a literal line with a regrain rule → the rule's own arithmetic over
  its month cells — ``=SUM(...)`` / ``=AVERAGE(...)`` / the last
  month's ref;
* a formula line → its OWN formula rendered at the bucket, operands
  resolving to their bucket cells (via :func:`project_model`, whose
  projected tree keys lines by the same qpaths) — exactly what a
  hand-built book does at the quarter column;
* a literal line with NO rule → the projection's ``#VALUE!`` teaching
  token, never an invented sum.

Native month columns stay the writer's ordinary cells — only their
PHYSICAL positions shift right past the inserted bucket columns, and
addresses shift with them, so every native formula follows for free.
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from ...core.calendar import GREGORIAN
from ...core.time import article, bucket_ranges


#: The kind of a sheet's OWN periods in a plan — its native grain's
#: columns, whatever that grain is. A bucket is keyed by its ruler's
#: name (``'month'``, ``'quarter'``, ``'half'``, ``'year'``), so a day
#: sheet may carry month buckets beside its days without the two keys
#: meeting (they used to: the native columns were spelled ``'month'``).
NATIVE = "period"

#: One interleaved column: ``(NATIVE, i)`` / ``('quarter', bi)`` /
#: ``('year', bi)`` — in on-sheet order.
ColKey = Tuple[str, int]


@dataclass
class TotalsPlan:
    """The interleave for ONE sheet."""

    kinds: Tuple[str, ...]
    n: int                                   # native period count
    #: on-sheet order of every period-area column
    columns: List[ColKey] = field(default_factory=list)
    #: ColKey -> physical 1-based column number
    col_of: Dict[ColKey, int] = field(default_factory=dict)
    #: kind -> [(bucket_label, lo, hi)] — half-open native ranges
    buckets: Dict[str, List[Tuple[str, int, int]]] = field(
        default_factory=dict
    )
    #: the sheet's label style ('finance' | 'iso' | 'compact')
    style: str = "finance"
    #: the word total columns wear on the month row ("Total", or the
    #: model's own word)
    totals_word: str = "Total"
    #: header depth: 1 + one row per bucket kind — "year" over
    #: "quarters" over "months". Data starts at ``depth + 1``.
    depth: int = 1
    #: Track column groups: how many columns one period-area column
    #: key occupies. ``col_of`` speaks GROUPED space — where a group
    #: STARTS — and a specific cell is ``cell(key, slot)``. Stride 1
    #: (every book without a compare layout) collapses both to the
    #: historical single-column reading.
    stride: int = 1
    #: The sheet's own ruler — what its native captions shorten to under
    #: a tier (an hour keeps its clock, "13:00", under its day).
    grain: Optional[str] = None


    def native_col(self, i: int) -> int:
        """The column of the sheet's own period ``i``."""
        return self.col_of[(NATIVE, i)]

    #: The name this had when every native period was a month.
    month_col = native_col

    def cell(self, key: ColKey, slot: int = 0) -> int:
        """The physical column of one SLOT inside ``key``'s group."""
        return self.col_of[key] + slot


def build_totals_plan(
    kinds: Tuple[str, ...],
    grain: str,
    start: Any,
    n: int,
    base_col: int,
    stride: int = 1,
) -> Optional[TotalsPlan]:
    """The interleave for a sheet whose periods start at ``base_col``.

    ``None`` when nothing is to be inserted — totals undeclared, no
    periods, or no declared kind is coarser than the native grain (a
    quarterly-native sheet still takes a year column; a yearly one
    takes nothing).
    """
    usable = tuple(
        k for k in kinds if _coarser(k, grain)
    )
    if not usable or n <= 0:
        return None
    plan = TotalsPlan(kinds=usable, n=n)
    for kind in usable:
        plan.buckets[kind] = [
            (label, lo, hi)
            for (label, lo, hi) in bucket_ranges(grain, kind, str(start), n)
        ]
    # Walk the periods; a bucket column follows its last period —
    # quarter before year where both close on the same month.
    ends: Dict[str, Dict[int, int]] = {
        kind: {hi: bi for bi, (_l, _lo, hi) in enumerate(plan.buckets[kind])}
        for kind in usable
    }
    for i in range(n):
        plan.columns.append((NATIVE, i))
        for kind in usable:
            bi = ends[kind].get(i + 1)
            if bi is not None:
                plan.columns.append((kind, bi))
    for offset, key in enumerate(plan.columns):
        plan.col_of[key] = base_col + offset * stride
    plan.depth = 1 + len(usable)
    plan.stride = stride
    plan.grain = grain
    return plan


def _coarser(kind: str, grain: str) -> bool:
    """Whether ``kind`` is a bucket a ``grain`` sheet can total by — a
    ruler ``grain`` nests in; an unknown word never is."""
    return GREGORIAN.coarser(kind, grain)


class TotalsWarning(UserWarning):
    """A view declares a totals kind a sheet cannot draw — its periods do
    not lie inside that kind's (weeks in months), or it straddles another
    declared kind's (weeks beside months over days). The column is left
    out; the warning says why."""


def drawable_kinds(kinds: Tuple[str, ...], grain: str) -> Tuple[Tuple[str, ...], List[str]]:
    """The declared totals ``kinds`` a ``grain`` sheet draws, in their
    order, and a note for each it cannot. A kind finer than the sheet
    (a month column on a quarterly sheet) is skipped without a word —
    one view serves sheets of several grains. A kind the sheet's periods
    do not lie inside (a month column over weeks) is a note; so is a
    kind that does not nest with one already kept (weeks and months over
    days — the tiers of a header must nest)."""
    kept: List[str] = []
    notes: List[str] = []
    up = GREGORIAN.rolls_up_to(grain) if GREGORIAN.knows(grain) else ()
    for kind in kinds:
        if kind in kept or not (GREGORIAN.knows(kind) and GREGORIAN.knows(grain)):
            continue
        if kind == grain or GREGORIAN.nests(kind, grain):
            continue
        if not GREGORIAN.nests(grain, kind):
            notes.append(
                f"'{kind}' totals are not drawn: a {grain} does not lie inside one "
                f"{kind} ({grain}s straddle {kind}s). A {grain} sheet totals by "
                f"{', '.join(up) if up else 'nothing coarser'}."
            )
            continue
        clash = next((k for k in kept if not GREGORIAN.comparable(k, kind)), None)
        if clash is not None:
            notes.append(
                f"'{kind}' totals are not drawn beside '{clash}' on {article(grain)} {grain} sheet: "
                f"{kind}s and {clash}s do not nest, so their header tiers cannot "
                f"stand one over the other. Total by one of them."
            )
            continue
        kept.append(kind)
    return tuple(kept), notes


def display_label(bucket_label: str, kind: str, style: str) -> str:
    """``2026-Q1`` → ``Q1 2026`` (finance) — same voice as the month
    labels beside it; iso keeps the native spelling. A month bucket (a
    day sheet's month total) is written as the month ruler writes a
    month: ``Jan 2026``, ``Jan-26``."""
    if style == "iso":
        return bucket_label
    if kind in ("day", "month", "week", "iso_year"):
        ruler = GREGORIAN.ruler(kind)
        try:
            return ruler.label(ruler.parse(bucket_label), style)
        except ValueError:
            return bucket_label
    tag = _TAG.get(kind)
    if tag and f"-{tag}" in bucket_label.upper():
        y, n = bucket_label.upper().split(f"-{tag}")
        return f"{tag}{n} {y}"
    return bucket_label


#: The letter a quarter's and a half's keys carry (``2026-Q1``,
#: ``2026-H2``) — what the finance spelling leads with (``Q1 2026``).
_TAG = {"quarter": "Q", "half": "H"}


def regrain_rule_of(var: Any) -> Any:
    """The line's declared coarsening rule, when the writer can spell it
    as Excel arithmetic: a recipe name (``'sum'``, ``'mean'``,
    ``'geometric'``, …) or a :class:`~modeleon.core.regrain.Ratio`.
    ``None`` otherwise (a callable rule, or no rule)."""
    from ...core.regrain import RECIPES, Ratio

    spec = getattr(var, "_regrain", None)
    rule = getattr(spec, "default", None) if spec is not None else None
    if isinstance(rule, str) and rule in RECIPES:
        return rule
    if isinstance(rule, Ratio):
        return rule
    return None


def _cell(ref: str) -> Tuple[int, int]:
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1
    col = 0
    for ch in ref[:i].upper():
        col = col * 26 + (ord(ch) - 64)
    return int(ref[i:]), col


def compact_refs(refs: List[str]) -> List[Tuple[str, Optional[str]]]:
    """Consecutive cells as ``(first, last)`` runs — along a row or down a
    column — and a lone cell as ``(cell, None)``. The cells of one bucket
    are NOT contiguous once subtotal columns stand between them, so a
    single ``first:last`` range would swallow the subtotals."""
    runs: List[List[str]] = []
    prev: Optional[Tuple[int, int]] = None
    step: Optional[Tuple[int, int]] = None
    for ref in refs:
        here = _cell(ref)
        if runs and prev is not None:
            d = (here[0] - prev[0], here[1] - prev[1])
            if d in ((0, 1), (1, 0)) and (step is None or d == step):
                runs[-1].append(ref)
                step, prev = d, here
                continue
        runs.append([ref])
        step, prev = None, here
    return [(r[0], r[-1] if len(r) > 1 else None) for r in runs]


def bucket_formula(
    recipe: str,
    refs: List[str],
    weights: Optional[List[int]] = None,
    qualify: Any = None,
    ranges: bool = True,
    blanks: bool = True,
) -> Optional[str]:
    """The Excel arithmetic (no leading ``=``) that folds one bucket of a
    line's native cells exactly as the engine's re-grain does
    (:func:`modeleon.core.regrain.apply_recipe`), a blank cell at zero:
    ``mean`` over the periods entered, weighted by their days when they
    differ (0 when none is entered), ``geometric`` compounded as
    ``EXP(LN(1+x)+…)-1`` — which is ``#NUM!`` where a rate is ``-100%``
    or worse, as the engine's value is. ``qualify`` prefixes a cell with
    its sheet when the bucket reads another sheet; ``ranges=False``
    lists every cell instead of joining runs into ranges; ``blanks=False``
    says every period of the bucket is entered, and a mean is then the
    plain weighted form. ``None`` for a recipe Excel arithmetic does not
    spell."""
    if not refs:
        return None
    q = qualify or (lambda ref: ref)
    if recipe == "first":
        return q(refs[0])
    if recipe == "last":
        return q(refs[-1])
    pieces = ([q(a) if b is None else f"{q(a)}:{b}" for a, b in compact_refs(refs)]
              if ranges else [q(r) for r in refs])
    if recipe == "geometric":
        logs = [f"SUMPRODUCT(LN(1+{p}))" if ":" in p else f"LN(1+{p})" for p in pieces]
        return "EXP(" + "+".join(logs) + ")-1"
    if recipe == "mean":
        weighted = bool(weights) and len(set(weights or ())) > 1
        terms = ("+".join(f"{q(r)}*{w}" for r, w in zip(refs, weights or ()))
                 if weighted else "")
        if not blanks:
            return (f"({terms})/{sum(weights or ())}" if weighted
                    else f"AVERAGE({','.join(pieces)})")
        # Over the periods entered - a blank one skipped, as AVERAGE skips
        # it - and 0 when none is: COUNTA, not IFERROR, so an error cell
        # still spreads, as the engine's bucket takes it.
        none_entered = f"COUNTA({','.join(pieces)})=0"
        if weighted:
            days = "+".join(f"ISNUMBER({q(r)})*{w}" for r, w in zip(refs, weights or ()))
            return f"IF({none_entered},0,({terms})/({days}))"
        return f"IF({none_entered},0,AVERAGE({','.join(pieces)}))"
    fn = {"sum": "SUM", "min": "MIN", "max": "MAX"}.get(recipe)
    if fn is None:
        return None
    return f"{fn}({','.join(pieces)})"


#: Excel's own limits on one formula.
_MAX_ARGS, _MAX_CHARS = 255, 8192


def listed_or_ranged(recipe: str, refs: List[str], weights: Optional[List[int]] = None,
                     qualify: Any = None, blanks: bool = True) -> Optional[str]:
    """Every cell listed, as a hand-built subtotal column reads - unless the
    list would break Excel's limits (a year of days), then runs as ranges."""
    body = bucket_formula(recipe, refs, weights, qualify, ranges=False, blanks=blanks)
    if body is not None and (len(refs) > _MAX_ARGS or len(body) > _MAX_CHARS - 1):
        body = bucket_formula(recipe, refs, weights, qualify, ranges=True, blanks=blanks)
    return body


def rule_formula(rule: str, refs: List[str],
                 weights: Optional[List[int]] = None, blanks: bool = True) -> Optional[str]:
    """The bucket cell for a rule-based line (see :func:`listed_or_ranged`)."""
    body = listed_or_ranged(rule, refs, weights, blanks=blanks)
    return None if body is None else f"={body}"


def short_label(label: str, kind: str, style: str, hierarchical: bool,
                grain: Optional[str] = None) -> str:
    """The label a HIERARCHICAL header wants: the year lives on its own
    row, so the rows below drop it — "Jan", "Q1" — while a flat header
    keeps the full form. Non-finance styles keep their native spelling
    either way. ``grain`` is the sheet's own ruler: a clock ruler's
    captions keep their clock ("13:00"), the day on the tier above."""
    if not hierarchical or style != "finance":
        return display_label(label, kind, style) if kind != NATIVE else label
    if kind == NATIVE:
        if grain is not None and GREGORIAN.knows(grain) and GREGORIAN.ruler(grain).sub_day:
            return label.rsplit(" ", 1)[-1] if " " in label else label
        return label.split(" ")[0] if " " in label else label
    if kind in ("day", "month", "week"):
        # A month or week tier over days: "Jan", "W05" — the year on
        # the tier above.
        full = display_label(label, kind, style)
        return full.split(" ")[0] if " " in full else full
    tag = _TAG.get(kind)
    if tag and f"-{tag}" in label.upper():
        return f"{tag}{label.upper().split(f'-{tag}')[1]}"
    return label


def header_rows(
    plan: TotalsPlan, month_labels: List[str]
) -> List[dict]:
    """The multi-row header — one list per row, top (coarsest) first:

        [{row: 1, cells: [{col, colspan, label, kind}]},   # years
         {row: 2, cells: [...]},                           # quarters
         {row: 3, cells: [...]}]                           # months

    A bucket's span covers its months and every finer bucket column
    inside it — but NOT its own total column, which stands beside the
    span carrying its own name ("Total Q1") from the bucket's OWN tier
    down through the month row. A quarter's total is a quarter-level
    number, so it starts where "Q1" starts; the year's starts where
    "2026" does. Plain data, not openpyxl calls: the .xlsx writer
    paints and merges it, and any other renderer can draw the same
    header from it.
    """
    # Coarsest tier on top — the calendar's order, whatever order the
    # kinds were declared in; a month tier stands over a day sheet's
    # days.
    kinds_top_down = sorted(plan.kinds, key=GREGORIAN.rank, reverse=True)
    rows: List[dict] = []
    hierarchical = True

    def tiered(kind: str) -> bool:
        # A day tier drops its month and year only under a month tier
        # that carries them; over an hour sheet alone it reads in full.
        return hierarchical and (kind != "day" or "month" in plan.kinds)

    for row_i, kind in enumerate(kinds_top_down, start=1):
        cells = []
        for bi, (blabel, lo, hi) in enumerate(plan.buckets[kind]):
            first = plan.col_of[(NATIVE, lo)]
            # …up to but NOT including its own total column: that one
            # carries the bucket's "Total" label, starting on this very
            # row. Swallowing it left the label nowhere to go but the
            # month row, one tier below the thing it totals. The total
            # is a GROUP too, so "before it" means one full stride.
            last = plan.col_of[(kind, bi)] - 1
            if last < first:
                continue
            cells.append({
                "col": first,
                "colspan": last - first + 1,
                "label": short_label(blabel, kind, plan.style, tiered(kind)),
                "kind": kind,
            })
        rows.append({"row": row_i, "cells": cells})
    month_row = len(kinds_top_down) + 1
    cells = [
        {
            "col": plan.native_col(i),
            "colspan": plan.stride,
            "label": short_label(
                month_labels[i] if i < len(month_labels) else "",
                NATIVE, plan.style, hierarchical, plan.grain,
            ),
            "kind": NATIVE,
        }
        for i in range(plan.n)
    ]
    rows.append({"row": month_row, "cells": cells})
    # The total columns get NAMED — "Total Q1", "Total 2026" — an
    # unlabeled bold column under a merged span reads as "what is
    # this". Each label starts on its OWN tier and merges DOWN through
    # the month row: "Total Q1" begins where "Q1" begins, "Total 2026"
    # where "2026" does. Starting one row lower put a quarter-level
    # number on the month tier, reading as a thirteenth month.
    for j, kind in enumerate(kinds_top_down or ["quarter"]):
        if kind not in plan.buckets:
            continue
        label_row = j + 1
        rowspan = month_row - label_row + 1
        target = next(r for r in rows if r["row"] == label_row)
        for bi, (blabel, _lo, _hi) in enumerate(plan.buckets[kind]):
            target["cells"].append({
                "col": plan.col_of[(kind, bi)],
                "colspan": plan.stride,
                "rowspan": rowspan,
                "label": (
                    f"{plan.totals_word} "
                    f"{short_label(blabel, kind, plan.style, tiered(kind))}"
                ),
                "kind": kind,
            })
        target["cells"].sort(key=lambda c: c["col"])
    return rows
