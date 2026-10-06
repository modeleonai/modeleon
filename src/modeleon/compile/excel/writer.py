# SPDX-License-Identifier: Apache-2.0
"""Excel writer — pipeline step 3 of 3 (layout → translator → writer).

Entry point ``to_excel(path, root=...)`` is called by
:meth:`MultiVariableBase.to_excel`. Takes an MV root (any MultiVariable),
collects tabs (first-depth sub-MVs or explicit ``excel_props={'tab': True}``
descendants), runs layout to get addresses, then walks each tab writing
cells: input Variables become values, computed Variables become Excel
formulas (``=B4*B5``) resolved by :class:`ExcelTranslator`.

Typical caller path::

    model = mo.MultiVariable("Forecast")
    model.pnl = mo.MultiVariable("P&L", excel_props={'tab': True})
    model.pnl.revenue = mo.Variable(1_000_000, display_name="Revenue")
    model.pnl.cogs = model.pnl.revenue * mo.Variable(0.6, display_name="COGS %")

    model.to_excel("forecast.xlsx")    # lands here via to_excel(path, root=model)
"""

from __future__ import annotations

import logging
import re
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.worksheet.worksheet import Worksheet

from .translator import ExcelTranslator
from .layout import LayoutEngine, effective_hidden, row_kind, sheet_name_for
from .addresses import VariableAddresses
from .view import resolve_excel_view
from ...core.expr import ListExpr, Literal
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable

logger = logging.getLogger(__name__)

_UNRESOLVED_PLACEHOLDER = re.compile(r"\{[A-Z_][A-Z0-9_]*\}")


def _refuse_pending_rows(root: "MultiVariableBase") -> None:
    """A row still waiting on a loop that has not closed holds a
    provisional value; a workbook written now would carry it as a fact."""
    from ...core.loops import pending_message
    for var in root._collect_variables().values():
        if var._awaits or var._forward_of is not None:
            raise ValueError(pending_message(var))


def to_excel(path: str | Path, root: "MultiVariableBase") -> None:
    """Write a model to an ``.xlsx`` file.

    Called via ``root.to_excel(path)`` — the inherited method on any
    ``MultiVariable``. Every sheet-roled MV reachable from ``root``
    becomes a tab in the workbook. When no explicit sheet markers exist,
    first-depth sub-MVs are promoted to tabs automatically.

    Returns ``None`` (matches pandas/polars writer convention). The
    caller already has the path they passed in.
    """
    path = Path(path)
    _refuse_pending_rows(root)
    from .view import resolve_excel_view
    track_mode = resolve_excel_view(root).tracks or 'rows'
    root = expand_tracked_tree(root, mode=track_mode)
    # (no compare SELECTION is passed here: a declared tracks='compare'
    # shows every declared track; ``expand_tracked_tree(tracks=...)``
    # is how a caller narrows it)
    roots = _collect_root_sheets(root)
    if not roots:
        raise ValueError(
            "Nothing to emit: this MultiVariable has no content to put in the "
            "workbook. Add Variables or nested MultiVariables first, e.g.\n"
            "    pnl = mo.MultiVariable('P&L', excel_props={'tab': True})\n"
            "    pnl.revenue = mo.Variable(1_000_000)\n"
            "Then call to_excel() on the enclosing MultiVariable."
        )

    # ``flatten_nested_sheets=True`` keeps nested ``_is_sheet=True``
    # MVs as visual sections inside their containing sheet — the same
    # positional rule the HTML repr applies, so a tree like
    # ``t._is_sheet=True; t.h.pnl._is_sheet=True`` writes ``pnl`` as a
    # section row in t's sheet rather than dropping it.
    engine = LayoutEngine(roots, flatten_nested_sheets=True, header_rows=True)
    addresses = engine.compute_addresses()
    if not addresses:
        raise ValueError(
            "Tabs exist but contain no Variables. "
            "Every tab needs at least one Variable so the writer has "
            "something to write. Example:\n"
            "    inputs = mo.MultiVariable('Inputs', excel_props={'tab': True})\n"
            "    inputs.revenue = mo.Variable(1_000_000, display_name='Revenue')"
        )

    # Resolve a UNIQUE, Excel-legal tab name per root ONCE so every layer
    # (worksheet title, ``var_to_sheet``, ``current_sheet``) shares the SAME
    # name. A model and its ``.at(grain=…)`` clone both derive ``'Saas'``; if the
    # dedup name didn't reach ``var_to_sheet`` the qualifier check would read a
    # genuinely cross-tab ref as same-sheet and drop the ``'Saas1'!`` prefix.
    # Because these names are already unique + ≤31 chars, ``create_sheet`` uses
    # them verbatim (no openpyxl re-dedup), so ``ws.title`` matches. See
    # ``_unique_sheet_names`` / ``_maybe_qualify``.
    sheet_names = _unique_sheet_names(roots)
    _refuse_lines_off_the_window(engine, roots, sheet_names, addresses)

    var_to_sheet = _build_var_to_sheet(addresses, roots, sheet_names)
    book_time = book_time_for(engine, roots, sheet_names, addresses, var_to_sheet)
    identity_cells = _identity_homes(roots, addresses)
    translator = ExcelTranslator(addresses, var_to_sheet,
                                 identity_cells=identity_cells, time=book_time)

    wb = Workbook()
    # ``Workbook()`` always creates a default sheet — the assert pins
    # that contract for type-checkers (Workbook.active is Optional).
    assert wb.active is not None
    wb.remove(wb.active)
    # The book's own font: the Normal style every cell starts from —
    # the empty ones, and the ones a reader types into later.
    book_font = set_book_font(wb, root)

    for sheet_mv, sheet_name in zip(roots, sheet_names):
        write_laid_out_sheet(
            wb.create_sheet(title=sheet_name), sheet_mv, engine, addresses,
            translator, sheet_name, var_to_sheet, book_font=book_font, time=book_time,
        )

    wb.save(path)


def set_book_font(wb: Workbook, root: "MultiVariableBase") -> dict:
    """Give the book ``root``'s font (its resolved view's ``font``) as the
    Normal style; returns that font in the ``excel_props`` vocabulary, for
    :func:`write_laid_out_sheet`. Public, with it, for an exporter that
    builds its own workbook."""
    book_font = _font_props(resolve_excel_view(root).base_font)
    _set_book_font(wb, book_font)
    return book_font


def write_laid_out_sheet(
    ws: Worksheet,
    sheet_mv: "MultiVariableBase",
    engine: LayoutEngine,
    addresses: dict[str, VariableAddresses],
    translator: ExcelTranslator,
    sheet_name: str,
    var_to_sheet: dict[str, str],
    *,
    book_font: Optional[dict] = None,
    time: Any = None,
) -> None:
    """Write one sheet of a laid-out book — every product of the layout
    handed to the sheet writer, then the subtotal, blend and track passes.

    The one door from a layout to a sheet: :func:`to_excel` and any
    exporter that builds its own workbook write each sheet through it, so
    a new product of the layout reaches every book at once.
    """
    # A helper tab (``excel_props={'hidden': True}``): the sheet is
    # written in full — values and formulas live, the audit still
    # works — but the tab is natively hidden.
    if effective_hidden(sheet_mv):
        ws.sheet_state = "hidden"
    key = id(sheet_mv)
    # The layout's key for this sheet — the tab title may differ
    # (``P/L`` is laid out as itself and written as ``P_L``).
    laid_out = engine.sheet_name_by_id.get(key, sheet_name)
    plan = engine.totals_by_sheet.get(laid_out)
    _write_sheet(ws, sheet_mv, addresses, translator, sheet_name,
                 engine.section_header_rows, engine.time_headers,
                 engine_meta=engine.meta_by_sheet,
                 board_var_ids=engine.board_var_ids,
                 section_header_echoes=engine.section_header_echoes,
                 inline_section_ids=engine.inline_section_ids,
                 period_start=engine.period_start_by_sheet.get(key, 0),
                 period_var_ids=engine.period_var_ids,
                 meta_fields=engine.meta_fields_by_sheet.get(key, ()),
                 constants_caption=engine.constants_by_sheet.get(key),
                 totals_plan=plan,
                 group_plan=engine.groups_by_sheet.get(laid_out),
                 row_outline=engine.row_outline_by_sheet.get(laid_out),
                 header=engine.header_by_sheet.get(key),
                 book_font=book_font,
                 columns=engine.columns_by_sheet.get(key),
                 top=engine.top_by_sheet.get(key, 0),
                 opening_caption=engine.opening_by_sheet.get(key),
                 freeze=engine.freeze_by_sheet.get(key),
                 articles=engine.articles,
                 caption_row=engine.caption_rows_by_sheet.get(key))
    if plan is not None:
        from .totals_writer import write_totals

        write_totals(ws, sheet_mv, plan, addresses, var_to_sheet, sheet_name)
    # The blend row references its track subrows instead of baking
    # the splice — fill an actual in the file and "live" moves.
    from .blend_writer import write_blend_links

    write_blend_links(ws, sheet_mv, addresses, var_to_sheet, sheet_name)
    # …and a DERIVED track subrow explains itself in its own
    # coordinates, instead of sitting as a literal under a row that
    # does. Runs after the blend pass: the blend row is its own
    # case and must not be overwritten here.
    from .track_writer import write_track_formulas

    write_track_formulas(ws, sheet_mv, addresses, var_to_sheet, sheet_name, time=time)


def _refuse_lines_off_the_window(
    engine: LayoutEngine,
    roots: list[MultiVariableBase],
    sheet_names: list[str],
    addresses: dict[str, VariableAddresses],
) -> None:
    """Stop the export when a line runs on a time window of its own.

    Every series on a sheet is written from the sheet's first period
    column, one cell per period of the sheet's window. A line declared
    with its own ``start=`` / ``grain=`` that differs from that window —
    another start, another grain, or another number of periods — would
    not line up with the sheet's dates, and formulas reading it would
    pair the wrong periods: the workbook would compute other numbers
    than Python does. Excel export cannot place such a line yet, so it
    refuses before anything is written. A line whose own window equals
    its sheet's, and a model with no window at all, are unaffected.
    """
    found = []
    for sheet_mv, sheet_name in zip(roots, sheet_names):
        window = engine.window_by_sheet.get(id(sheet_mv))
        if window is None or window.start is None or window.grain is None:
            continue    # undated columns: no date to put a line under
        for var in engine._iter_sheet_variables(sheet_mv):
            if var.id not in addresses:
                continue
            source, loc = var, _off_window(var, window)
            if loc is None and var._expr is not None:
                read = _reads_unwritten_off_window(
                    var._expr, window, addresses, set()
                )
                if read is not None:
                    source, loc = read, read.time
            if loc is not None:
                found.append((var, source, loc, sheet_name, window))
    if not found:
        return

    def _name(v: Variable) -> str:
        return (f"'{v.display_name}'" if v.path.is_floating
                else f"'{v.display_name}' ({v.path})")

    def _span(start: object, grain: object, n: Optional[int], first_day: Any = None) -> str:
        periods = "" if n is None else f", {n} period(s)"
        when = f"{first_day.isoformat()} (inside {start})" if first_day else start
        return f"starts {when}, grain '{grain}'{periods}"

    shown = []
    for var, source, loc, sheet_name, window in found[:10]:
        reads = "" if source is var else f" reads {_name(source)}, which"
        shown.append(
            f"  - {_name(var)}{reads} "
            f"{_span(loc.start, loc.grain, _periods_of(source), getattr(loc, 'first_day', None))}; "
            f"sheet '{sheet_name}' "
            f"{_span(window.start, window.grain, window.periods, getattr(window, 'first_day', None))}"
        )
    if len(found) > 10:
        shown.append(f"  - … and {len(found) - 10} more")
    raise ValueError(
        "Excel export does not support a line on a time window of its "
        "own yet. Every series is written from its sheet's first period "
        "column, one cell per period of the sheet's window, so these "
        "would not line up with the sheet's dates and formulas reading "
        "them would compute other numbers than Python. Nothing was "
        "written.\n" + "\n".join(shown) + "\n"
        "Put each such line on its sheet's window instead: drop its "
        "start= and grain=, and give it one value per period of that "
        "window, padding the periods it does not cover (e.g. with "
        "zeros). A series that starts with the window and ends early "
        "may keep its shorter list if it declares extend=mo.zero() or "
        "extend=mo.hold()."
    )


def _periods_of(var: Variable) -> Optional[int]:
    value = var._value
    if isinstance(value, list):
        return len(value)
    return getattr(value, 'time_length', None)


def _off_window(var: Variable, window: Any) -> Any:
    """``var``'s time location when it is a series that does not run
    on ``window``; ``None`` when it does, or carries no dates at all (a
    constant, a coordinate row, a row of axis labels)."""
    from ...core.time import Time, _parse
    from .layout import _is_periodic

    if isinstance(var, Time) or not _is_periodic(var):
        return None
    loc = var.time
    if loc is None or loc.start is None:
        return None
    if loc.grain != window.grain:
        return loc
    if getattr(loc, 'first_day', None) != getattr(window, 'first_day', None):
        # The same label, another first day: its first column would hold
        # a whole period where the sheet's is short, or the reverse.
        return loc
    if loc.start != window.start:
        try:
            # '2026-1' and '2026-01' name the same month.
            if _parse(loc.start, loc.grain) != _parse(window.start, window.grain):
                return loc
        except ValueError:
            return loc
    # A line placed by its own start= / grain= keeps exactly the values
    # it was given — it is never stretched to the window the way a bare
    # list with extend= is. Shorter or longer than the window, a formula
    # reading it in the book would pair periods Python does not pair.
    if var._grain is not None and window.periods is not None:
        n = _periods_of(var)
        if n is not None and n != window.periods:
            return loc
    return None


def _reads_unwritten_off_window(
    expr, window: Any, addresses: dict, seen: set,
) -> Optional[Variable]:
    """The first input ``expr`` reads that the workbook does not write
    and that is declared on a window of its own, or ``None``.

    A line the book writes is checked where it is written. One it does
    not write (a Variable left outside the model) is inlined into the
    formula as a literal list starting at the sheet's first period, so
    a series dated otherwise lands under the wrong dates. Anonymous
    intermediates (``a + b`` inside a longer formula) have no cells
    either — the formula is spelled through them, so the walk goes
    through them too, on a stack of its own: ``sum(rows)`` is one such
    step per row.
    """
    pending = [iter(expr.iter_refs())]
    while pending:
        ref = next(pending[-1], None)
        if ref is None:
            pending.pop()
            continue
        if id(ref) in seen or ref.id in addresses:
            continue
        seen.add(id(ref))
        if ref._expr is not None:
            pending.append(iter(ref._expr.iter_refs()))
        elif ref._grain is not None and _off_window(ref, window) is not None:
            return ref
    return None


def _has_tracked(mv: "MultiVariableBase") -> bool:
    from ...core.tracks import TrackValues
    for child in mv._components.values():
        if isinstance(child, MultiVariableBase):
            if _has_tracked(child):
                return True
        elif isinstance(getattr(child, '_value', None), TrackValues):
            return True
    return False


#: Metric columns of ``tracks='compare'``, in REGISTRY order — one
#: fixed catalogue, so every renderer agrees on what a deviation means.
#: Values are FRACTIONS wearing a percent number format (Excel's own
#: semantics); a renderer that prints plain percentages multiplies
#: the same numbers by 100.
_METRIC_REGISTRY = ("var", "var%", "share", "growth%")
_METRIC_LABELS = {
    "var": "Var", "var%": "Var %", "share": "Share", "growth%": "Growth %",
}
_METRIC_FORMATS = {"var%": "0.0%", "share": "0.0%", "growth%": "0%"}
#: Which metrics need the given/follow PAIR visible — without it the
#: number has no meaning and the column drops.
_METRIC_NEEDS_PAIR = {"var", "var%"}
#: Other spellings of the same metrics, accepted as selections; a
#: column selected this way is captioned the way it was spelled.
_METRIC_SPELLINGS = {
    "откл": ("var", "откл"),
    "откл%": ("var%", "откл %"),
    "уд.вес": ("share", "уд. вес"),
    "%роста": ("growth%", "% роста"),
}
#: The line a share is measured against, in either spelling.
_TOTAL_NAMES = ("total", "итого")
#: The track a deviation is taken from when the model has one so named.
_PLAN_NAMES = ("plan", "план")


def _resolve_metric_keys(selection, pair_available) -> list[tuple[str, str]]:
    """The metric selection as ``(key, caption)`` pairs: ``None`` = the
    deviation alone, ``[]`` = explicitly none; registry order
    regardless of the order the selection lists them in."""
    chosen: dict[str, str] = {}
    for spelled in (["var"] if selection is None else selection):
        key, caption = _METRIC_SPELLINGS.get(
            spelled, (spelled, _METRIC_LABELS.get(spelled, spelled)))
        chosen.setdefault(key, caption)
    return [
        (k, chosen[k]) for k in _METRIC_REGISTRY
        if k in chosen and (k not in _METRIC_NEEDS_PAIR or pair_available)
    ]


def _section_total_of(child):
    """The denominator line for a share: the nearest ``total`` sibling
    at or above this line's section, found by walking up the ownership
    chain. ``None`` when no ancestor section carries one (the metric
    stays honestly empty)."""
    from ...core.multi_variable import MultiVariableBase

    name = getattr(child, '_name_in_parent', None)
    node = getattr(child, '_owner', None) or getattr(child, '_parent', None)
    # A line NAMED total measures against the section above its own.
    if name in _TOTAL_NAMES and node is not None:
        node = (getattr(node, '_parent', None)
                or getattr(node, '_owner', None))
    while node is not None:
        if isinstance(node, MultiVariableBase):
            total = next((node._components[n] for n in _TOTAL_NAMES
                          if n in node._components), None)
            if (total is not None and total is not child
                    and isinstance(total, Variable)):
                return total
        node = (getattr(node, '_parent', None)
                or getattr(node, '_owner', None))
    return None


class _ComparePlan:
    """The column plan of ONE track group — shared by every line on
    the sheet, because it is derived from the DECLARATION, never from
    a line's own roles: a line may author a subset, and a plan read
    off whichever line came first would misalign the columns."""

    __slots__ = (
        "slot_of", "hidden", "given", "follow",
        "metric_keys", "metric_base", "stride", "pair",
    )

    def __init__(self, slot_of, hidden, given, follow,
                 metric_keys, metric_base, stride, pair):
        #: role → slot. Visible tracks first (given, follow, then the
        #: rest of the selection), metric columns next (registry order,
        #: slots ``metric_base + i``), hidden blend operands last — so
        #: the VISIBLE part of a group reads as the selection asked,
        #: and the file-only extras trail it.
        self.slot_of = slot_of
        self.hidden = hidden               # set of hidden ROLE names
        self.given = given
        self.follow = follow
        self.metric_keys = metric_keys
        self.metric_base = metric_base
        self.stride = stride
        #: (minuend, subtrahend) for the pair metrics — ANY two
        #: visible columns, by one generalized rule: the
        #: subtrahend is the plan-like column (a track named plan, else
        #: the follow), the minuend actual-like (given, else the blend,
        #: else the first other column). Tying the pair to the blend's
        #: given/follow would drop the deviation from a plain
        #: actual+plan comparison.
        self.pair = pair


def _compare_plan(decl, selection, roles,
                  metrics=None) -> Optional[_ComparePlan]:
    """Slots for one period group, in this order: given, follow, then
    the rest of the selection; then the METRIC columns (deviation,
    share, …); the blend and both its sources always ride (hidden
    when not selected) so the splice formula keeps its operands. Pair
    metrics need both blend tracks SELECTED — a deviation between
    anything else is a number without a meaning."""
    spec = getattr(decl, 'blend', None) if decl is not None else None
    given = getattr(spec, 'given', None)
    follow = getattr(spec, 'follow', None)
    blend = getattr(spec, 'name', None)
    declared = list(decl.names) if decl is not None else list(roles)
    all_roles = declared + ([blend] if blend and blend not in declared else [])
    if not all_roles:
        return None
    checked = [t for t in (selection or all_roles) if t in all_roles]
    if not checked:
        checked = all_roles
    visible: list[str] = []
    for r in (given, follow):
        if r and r in checked:
            visible.append(r)
    for r in checked:
        if r not in visible:
            visible.append(r)
    pair = None
    if len(visible) >= 2:
        plan_like = next(
            (r for r in _PLAN_NAMES if r in visible and r != given), None)
        sub = plan_like or (follow if follow in visible else visible[1])
        minuend = (
            given if given in visible and given != sub
            else blend if blend in visible and blend != sub
            else next((n for n in visible if n != sub), None)
        )
        if minuend and minuend != sub:
            pair = (minuend, sub)
    metric_keys = _resolve_metric_keys(metrics, pair is not None)
    metric_base = len(visible)
    hidden = [
        r for r in (given, follow, blend)
        if r and r in all_roles and r not in visible
    ]
    slot_of = {r: i for i, r in enumerate(visible)}
    for j, r in enumerate(hidden):
        slot_of[r] = metric_base + len(metric_keys) + j
    return _ComparePlan(
        slot_of, set(hidden), given, follow,
        metric_keys, metric_base,
        len(visible) + len(metric_keys) + len(hidden),
        pair,
    )


def _metric_variable(key, child, default_role, value, copies, plan, decl):
    """One metric COLUMN as a live Variable — one of the four in the
    registry, each an expression the file can hold:

    * ``var``     = given − follow (engine arithmetic, sums honestly);
    * ``var%``    = (given − follow) / ABS(follow), a FRACTION wearing a
      percent format — Excel's own percent semantics;
    * ``share``   = the line's shown series over the nearest section
      ``total``'s — AFTER aggregation by construction, because both
      operands' bucket cells are already aggregates;
    * ``growth%`` = shown over the same cell a year earlier, cell by
      cell (a ListExpr: early periods have no prior year and stay
      honestly EMPTY).

    ``None`` when the metric's ingredients are missing on this line —
    the column simply stays empty there, never a zero."""
    from ...core.expr import BinOp, FuncCall, Literal, ListExpr, Subscript, VarRef

    def _div(a, b):
        return (a / b) if (a is not None and b not in (None, 0)) else None

    shown_vals = value[default_role]
    n = len(shown_vals)

    if key in ("var", "var%"):
        if plan.pair is None:
            return None
        g = copies.get(plan.pair[0])
        f = copies.get(plan.pair[1])
        if g is None or f is None:
            return None
        if key == "var":
            dev = g - f
            return dev
        vals = [
            _div(
                (gv - fv) if (gv is not None and fv is not None) else None,
                abs(fv) if fv else None,
            )
            for gv, fv in zip(g._value, f._value)
        ]
        out = Variable(display_name='')
        out._value = vals
        out.var_type = 'list'
        out._set_expr(BinOp(
            op='/',
            left=BinOp(op='-', left=VarRef(var=g), right=VarRef(var=f)),
            right=FuncCall(func='ABS', args=[VarRef(var=f)]),
        ))
        return out

    if key == "share":
        total = _section_total_of(child)
        if total is None:
            return None
        tv = getattr(total, '_value', None)
        if type(tv).__name__ == 'TrackValues':
            troles = list(tv.roles)
            tdefault = (default_role if default_role in troles
                        else troles[0])
            tvals = tv[tdefault]
        elif isinstance(tv, list):
            tvals = tv
        else:
            return None
        if not isinstance(tvals, list):
            return None
        vals = [
            _div(sv, tvals[i] if i < len(tvals) else None)
            for i, sv in enumerate(shown_vals)
        ]
        out = Variable(display_name='')
        out._value = vals
        out.var_type = 'list'
        # VarRefs to the ORIGINALS: their ids equal the expanded
        # heads' ids (same component names), so the translator lands
        # on the head columns — the shown series of each line.
        out._set_expr(BinOp(
            op='/', left=VarRef(var=child), right=VarRef(var=total),
        ))
        return out

    if key == "growth%":
        from ...core.time import resolve_default_window
        window = resolve_default_window(child)
        grain = getattr(window, 'grain', None) or 'month'
        steps = {'month': 12, 'quarter': 4, 'year': 1}.get(grain)
        if not steps or n <= steps:
            return None
        # A short first period is no base for a year's growth: the period
        # a year after it would grow against part of a year.
        first = steps + 1 if getattr(window, 'first_day', None) is not None else steps
        items = []
        for t in range(n):
            if t < first:
                items.append(Literal(value=None))
                continue
            items.append(BinOp(
                op='/',
                left=Subscript(base=VarRef(var=child), key=t),
                right=Subscript(base=VarRef(var=child), key=t - steps),
            ))
        vals = [
            None if t < first else _div(shown_vals[t], shown_vals[t - steps])
            for t in range(n)
        ]
        out = Variable(display_name='')
        out._value = vals
        out.var_type = 'list'
        out._set_expr(ListExpr(items=items))
        # No prior-year operand is built for a BUCKET (a quarter or
        # year total) — the totals writer leaves these cells empty instead of
        # projecting a lag whose step no longer matches the grain.
        out._metric_no_bucket = True
        return out

    return None


def expand_tracked_tree(root: "MultiVariableBase",
                        mode: str = 'rows',
                        tracks: Optional[list[str]] = None,
                        metrics: Optional[list[str]] = None,
                        ) -> "MultiVariableBase":
    """Expand tracked Variables into per-track rows for emission.

    ``mode`` is the view's ``tracks`` setting (``ExcelView.tracks``):
    ``'rows'`` (default) — the display-default row plus labeled
    per-track rows; ``'blend'`` — the display-default row ONLY (the
    clean board book); ``'compare'`` — period-major column GROUPS: one
    column per track under each period, side by side.
    ``tracks`` narrows compare to a selection of track names;
    the blend and both its source tracks are always emitted so the
    splice formula stays honest — unselected ones ride as HIDDEN
    columns rather than silently degrading the blend to one operand.
    An all-scalar tracked line (a single-value field) stays a single
    cell in compare mode: a scalar has no period columns to group.

    The BLEND default sheet: a tracked line keeps its NAME and
    qpath on a row carrying its display-default series — the
    synthesized live track when the model blends, else the first
    declared track — so every formula that references the line renders
    against that row, and the file computes on the same blended series
    Python does. The other tracks follow as sibling value rows labeled
    "name · track-label".

    Untracked models return ``root`` unchanged — byte-identical output.
    Identity rides the project_model pattern: same python_name → same
    crystallized qpaths, so the translator resolves references in the
    expanded tree exactly as in the original.
    """
    from ...core.multi_variable import MultiVariable
    from ...core.tracks import TrackValues
    from ...core.tracks_decl import resolve_tracks_decl

    # 'compare' builds the same objects 'rows' does — the head line
    # carrying the display default plus one flat copy per other track
    # — but stamps each with its column SLOT instead of giving it a
    # row: the layout collapses a group onto one row at slot-offset
    # columns, and every period becomes a group of ``stride`` columns.
    if not _has_tracked(root):
        return root

    def _flat_copy(var: Variable, series: object) -> Variable:
        flat = Variable(display_name=var._display_name)
        flat._value = (list(series) if isinstance(series, list) else series)
        flat.var_type = 'list' if isinstance(series, list) else 'scalar'
        flat._unit = var._unit
        flat._start = var._start
        flat._grain = var._grain
        # The coarsening rule MUST ride the emission copy: the totals
        # writer spells a ruled line's bucket as its own arithmetic
        # (=SUM over the month cells). Dropped, the bucket fell to the
        # formula-translate path, whose cross-sheet operands inline
        # their VALUES — and a tracked operand's value is a TrackValues
        # container, which landed in the cell as machine repr.
        flat._regrain = var._regrain
        flat._track_copy = True     # a row of a line with tracks
        if hasattr(getattr(var, '_value', None), 'roles'):
            # A track's row folds a hole as its whole line does (a line of
            # dates keeps it blank even on a track with nothing entered).
            from ...core.regrain import hole_value
            flat._hole_value = hole_value(var._value)
        flat._excel_props = dict(var._excel_props)
        src_code = getattr(var, '_source_code', None)
        if src_code:
            flat._source_code = src_code
        return flat

    def _adopt_ref(out: "MultiVariableBase", name: str,
                   child: object) -> None:
        # Carry BY REFERENCE, bypassing ``__setattr__`` adoption: the
        # child stays owned by the ORIGINAL tree. Re-adoption would
        # reference-clone an already-owned node (cross-container
        # semantics) — self-referential ``=B2`` formulas and lost
        # display names in the emitted book. The expanded tree is a
        # throwaway emission view; shared children keep their original
        # identity, values, and view cascade. Writes hit the real
        # storage (``__dict__`` + ``_component_names``) — the
        # ``_components`` / ``_component_order`` accessors are derived
        # copies.
        out.__dict__[name] = child
        out.__dict__.setdefault('_component_names', []).append(name)

    def _walk(mv: "MultiVariableBase") -> "MultiVariableBase":
        out = MultiVariable(
            mv._display_name, excel_props=dict(mv._excel_props) or None
        )
        for name, child in mv._components.items():
            if isinstance(child, MultiVariableBase):
                from .view import ExcelView

                if isinstance(child, ExcelView):
                    # Presentation node — reference, never content.
                    _adopt_ref(out, name, child)
                    continue
                if not _has_tracked(child):
                    # Nothing to expand below — the whole subtree
                    # rides by reference, formulas intact.
                    _adopt_ref(out, name, child)
                    continue
                setattr(out, name, _walk(child))
                continue
            value = getattr(child, '_value', None)
            if not isinstance(value, TrackValues):
                _adopt_ref(out, name, child)
                continue
            decl = resolve_tracks_decl(child)
            blend_name = getattr(getattr(decl, 'blend', None), 'name', None)
            roles = list(value.roles)
            default_role = (blend_name if blend_name in roles else roles[0])
            default = _flat_copy(child, value[default_role])
            # Is this row's series a SPLICE of its own subrows, or a
            # function of its operands' blends? The engine's own
            # predicate (``_synthesize_blend``): a line with authored
            # roles — or none at all — is a DATA line and got spliced;
            # anything else inherited live through the broadcast. The
            # flat copy loses ``_role_kwargs``, so the answer travels
            # as a mark for the blend writer to read.
            if ((getattr(child, '_role_kwargs', None) is not None
                    or child._expr is None)
                    and not getattr(child, '_regrained_tracks', False)):
                default._blend_is_splice = True
            if default_role in (getattr(child, '_role_kwargs', None) or {}):
                # The shown series was entered, not derived — mark it
                # so consumers can distinguish the two.
                default._authored_track = default_role
            if child._expr is not None:
                # The shared formula renders against display-default
                # rows of its operands — Excel formulas compute on the
                # blended series, exactly the BLEND sheet's rule.
                default._set_expr(child._expr)
                # A running total folds as the running total of its
                # folded input, as the line itself does.
                if getattr(child, '_cumsum_source', None) is not None:
                    default._cumsum_source = child._cumsum_source
            else:
                # A track AUTHORED with an expression (``forecast=recurrence``,
                # ``plan=<row>``) keeps its AST: the shown row is that
                # track's formula, not a copy of its numbers. It is a
                # COORDINATE's formula — rendered inside that track
                # (operands resolve to their same-track rows), and never
                # borrowed by the track writer as the line's shared one.
                shown = child.shown_track()
                if shown is not None:
                    shown_role, track_expr = shown
                    default._set_expr(track_expr)
                    default._track_role = shown_role
                    default._expr_is_track_specific = True
            setattr(out, name, default)
            if mode == 'blend':
                continue    # the clean board book: default rows only
            if mode == 'compare':
                if not isinstance(value[default_role], list):
                    continue    # a scalar has no period columns to group
                plan = _compare_plan(decl, tracks, roles, metrics)
                if plan is None:
                    continue    # no declaration to group by
                default._col_slot = plan.slot_of.get(default_role, 0)
                copies: dict[str, Variable] = {default_role: default}
                for role, slot in plan.slot_of.items():
                    if role == default_role or role not in roles:
                        continue
                    extra = _flat_copy(child, value[role])
                    if role in (getattr(child, '_role_kwargs', None) or {}):
                        extra._authored_track = role
                    track_expr = child.track_expr(role)
                    if track_expr is not None:
                        extra._set_expr(track_expr)
                        extra._track_role = role
                    extra._col_slot = slot
                    extra._col_head = default
                    if role in plan.hidden:
                        extra._col_hidden = True
                    extra._col_word = (
                        decl.label(role) if (decl and role in decl) else role
                    )
                    copies[role] = extra
                    setattr(out, f"{name}__track_{role}", extra)
                default._col_word = (
                    decl.label(default_role)
                    if (decl and default_role in decl) else default_role
                )
                for mi, (mkey, caption) in enumerate(plan.metric_keys):
                    mvar = _metric_variable(
                        mkey, child, default_role, value,
                        copies, plan, decl,
                    )
                    if mvar is None:
                        continue
                    mvar._display_name = (
                        f"{child._display_name or name} · {caption}"
                    )
                    mvar._col_slot = plan.metric_base + mi
                    mvar._col_head = default
                    mvar._col_word = caption
                    fmt = _METRIC_FORMATS.get(mkey)
                    if mkey == 'var':
                        # The deviation is a difference of two values of
                        # the line, in its units: it wears the line's
                        # format, unless that format cannot print a
                        # signed number.
                        # General is written explicitly: a section's or
                        # the view's format would otherwise reach the
                        # column through inheritance all the same.
                        line_fmt = resolve_number_format(child)
                        if line_fmt:
                            fmt = (line_fmt if format_shows_signed_numbers(line_fmt)
                                   else 'General')
                    # The column is a cell of the line's row: it wears
                    # the line's look — its rule, fill and weight run on
                    # through the deviation column instead of breaking there.
                    line_look = {
                        k: v for k, v in row_look(child).items()
                        if k in _CELL_LOOK_KEYS
                    }
                    if line_look:
                        mvar._excel_props = {
                            **line_look, **(mvar._excel_props or {}),
                        }
                    if fmt:
                        mvar._excel_props = {
                            **(mvar._excel_props or {}),
                            'number_format': fmt,
                        }
                    setattr(out, f"{name}__metric_{mi}", mvar)
                continue
            for role in roles:
                if role == default_role:
                    continue
                extra = _flat_copy(child, value[role])
                if role in (getattr(child, '_role_kwargs', None) or {}):
                    # The subrow's series was entered, not derived —
                    # mark it so consumers can distinguish the two.
                    extra._authored_track = role
                track_expr = child.track_expr(role)
                if track_expr is not None:
                    # An authored-expression track is a formula row,
                    # rendered inside its own coordinate.
                    extra._set_expr(track_expr)
                    extra._track_role = role
                label = decl.label(role) if (decl and role in decl) else role
                extra._display_name = (
                    f"{child._display_name or name} · {label}"
                )
                setattr(out, f"{name}__track_{role}", extra)
        # Window + identity ride along so the timeline header and the
        # crystallized qpaths match the original tree.
        for w in ('default_grain', 'default_start', 'default_periods'):
            if getattr(mv, w, None) is not None:
                setattr(out, w, getattr(mv, w))
        # …with the Variable the window starts at, a short first period,
        # and the workbook header (untracked rows are adopted by
        # reference, so both still point at rows of the rebuilt tree).
        for slot in ('_default_start_source', '_default_first_day', 'default_header'):
            if slot in mv.__dict__:
                out.__dict__[slot] = mv.__dict__[slot]
        # The presentation slot rides too — subrows must resolve the
        # same view cascade (number formats, bands) as the original
        # tree, and the slot assignment is adoption-exempt so this is
        # a plain attribute copy.
        _dev = mv.__dict__.get('default_excel_view')
        if _dev is not None:
            out.default_excel_view = _dev
        decl = getattr(mv, 'tracks', None)
        if decl is not None:
            out.tracks = decl
        if mv._python_name is not None:
            out._python_name = mv._python_name
        return out

    return _walk(root)


def book_time_for(engine: Any, roots: list, sheet_names: list,
                  addresses: dict, var_to_sheet: dict) -> Any:
    """The workbook's time (see :mod:`.header`), with every header
    mirror registered on its sheet so references to it qualify."""
    from .header import build_book_time
    for sheet_mv, sheet_name in zip(roots, sheet_names):
        plan = engine.header_by_sheet.get(id(sheet_mv))
        if plan is not None:
            for _src, mirror in plan.mirrors:
                var_to_sheet[mirror.id] = sheet_name
    return build_book_time(engine, roots, sheet_names, addresses)


def _collect_root_sheets(root: "MultiVariableBase") -> list[MultiVariableBase]:
    """Collect sheets to emit from an explicit root.

    The root is the workbook itself — it never becomes its own tab.
    First-depth structure decides the sheet list:

    - First-depth sub-MVs → one tab each (using the user's MV when it's
      already sheet-roled, otherwise wrapped in a virtual sheet).
    - Direct Variables on ``root`` → collected into a single synthesized
      "overview" tab named after ``root`` so orphan inputs (no enclosing
      MV) still land in the workbook.

    Deeper sub-MVs (whether marked ``_is_sheet=True`` or not) render as
    sections inside their containing tab — :class:`LayoutEngine` is
    invoked with ``flatten_nested_sheets=True`` in :func:`to_excel`.

    **Flat-container exception.** A container can declare that its
    children are SECTIONS, not tabs, by carrying the duck-typed
    ``_children_are_sections`` marker. The first-depth promotion rule
    is then suppressed: the whole root collapses into a single virtual
    sheet and every nested MV renders as a section inside it.

    Note this is a statement about the container's CHILDREN, which is
    why ``excel_props={'tab': True}`` cannot express it — that says
    "I am a tab" and leaves each child a tab of its own. A container of
    like records is the case that needs it: a roster of people would
    otherwise open one tab per person, when the thing the author
    declared was one list.

    A child that asks for a tab OF ITS OWN still gets one — the marker
    sets the default for the container's children, it does not overrule
    a child that declared otherwise. So one landmark record can stand
    apart while the rest stay a list.

    ``_is_context`` is accepted as a synonym.
    """
    if getattr(root, "_children_are_sections", False) or getattr(
        root, "_is_context", False
    ):
        label = root.display_name or root.python_name or "Sheet1"
        items = [(n, root._components[n]) for n in root._component_order]
        claimed = [
            comp
            for _n, comp in items
            if isinstance(comp, MultiVariableBase) and comp._is_sheet
        ]
        sections = [
            (n, comp)
            for n, comp in items
            if not (isinstance(comp, MultiVariableBase) and comp._is_sheet)
        ]
        out: list[MultiVariableBase] = []
        if sections:
            out.append(_make_virtual_sheet(label, sections, source=root))
        out.extend(claimed)
        return out

    direct_vars: list[tuple[str, Variable]] = []
    sub_mvs: list[MultiVariableBase] = []
    for name in root._component_order:
        comp = root._components[name]
        if isinstance(comp, Variable):
            direct_vars.append((name, comp))
        elif isinstance(comp, MultiVariableBase):
            from .view import ExcelView

            if isinstance(comp, ExcelView):
                # A named view is presentation metadata — never a tab.
                continue
            sub_mvs.append(comp)

    out: list[MultiVariableBase] = []
    if direct_vars:
        overview_name = root.display_name or root.python_name or "Sheet1"
        # Propagate ``_excel_layout`` from the root — without this the
        # virtual sheet loses the user's layout overrides (sheet name,
        # format) when the root is a Variables-only Model.
        out.append(_make_virtual_sheet(overview_name, direct_vars, source=root))

    for mv in sub_mvs:
        if mv._is_sheet:
            out.append(mv)
        else:
            # Promote a plain MV to a virtual Sheet without mutating the
            # user's model — new wrapper, same components by reference.
            label = mv.display_name or mv.python_name or mv.id
            items = [(n, mv._components[n]) for n in mv._component_order]
            out.append(_make_virtual_sheet(label, items, source=mv))
    return out


def _make_virtual_sheet(
    name: str,
    components: list[tuple[str, Any]],
    source: MultiVariableBase | None = None,
) -> MultiVariableBase:
    """Bare-bones Sheet-roled MV for layout / writing.

    Uses ``object.__new__`` so the synthesized sheet doesn't attach
    itself to an enclosing ``with`` context or mint a QPath — it's a
    render-time view, not a real model node.

    ``source``, when given, is the user-authored MV the virtual sheet
    stands in for. We forward its ``_excel_layout`` so layout overrides
    (sheet name, format) declared on the original survive the wrap.
    Without this, ``mo.Model("forecast", excel_layout=...)`` would lose
    the override the moment the writer wraps its direct Variables.
    """
    from ...core.multi_variable import MultiVariable
    virt = object.__new__(MultiVariable)
    virt._role = 'sheet'
    virt._display_name = name
    virt._parent = None
    virt._name_in_parent = None
    virt._qualified_id = None
    virt._python_name = None
    # Carry the source's LAYOUT-BEARING props through the wrap. The
    # wrapper used to declare only ``tab``, which silently dropped an
    # ``orient`` the author had set — a records-oriented container
    # reached the layout as a plain MultiVariable and
    # was stacked section-per-record instead of laid out as the table
    # it declares itself to be. Styling props stay on the source's
    # cells; only orientation describes the SHEET.
    virt._excel_props = {'tab': True}
    _src_props = getattr(source, "_excel_props", None) or {}
    if _src_props.get('orient'):
        virt._excel_props['orient'] = _src_props['orient']
    # Hidden-ness travels with the source: a hidden source wrapped in a
    # virtual tab stays a HIDDEN sheet of the workbook.
    if effective_hidden(source):
        virt._excel_props['hidden'] = True
    virt._excel_layout = getattr(source, "_excel_layout", None) if source else None
    # The header cascade (``default_header``) starts at the MV the sheet
    # stands for — the wrapper has no parent to walk up from.
    virt._sheet_source = source
    # The view cascade must see THROUGH the wrapper: point the slot at
    # the source's resolved chain so a model-level ``default_excel_view``
    # (fonts, bands, formats) styles the virtual sheet too. Plain
    # attribute — the slot name is adoption-exempt.
    _src_view = None
    node = source
    while node is not None and _src_view is None:
        _src_view = node.__dict__.get('default_excel_view')
        node = getattr(node, '_parent', None) or getattr(node, '_owner', None)
    if _src_view is not None:
        virt.default_excel_view = _src_view
    # The TRACKS declaration rides through for the same reason: the
    # blend writer and the totals writer ask the SHEET which tracks it
    # splices and which one it displays, and a wrapper with no parent
    # could not answer. A records container is wrapped this way; without
    # the declaration every tracked line on a records sheet would show
    # the base track where Python computes the blend.
    from ...core.tracks_decl import resolve_tracks_decl
    _src_decl = resolve_tracks_decl(source) if source is not None else None
    if _src_decl is not None:
        virt.tracks = _src_decl
        # …and the WINDOW with it: the splice asks the sheet where the
        # close boundary falls, and an unanswerable window would read as
        # "no boundary at all" — every cell would take the
        # before-boundary form and the periods AFTER the close would
        # splice backwards.
        # Only the resolved values travel, so the wrapper answers
        # exactly as the source's chain would have.
        for _w in ('default_grain', 'default_start', 'default_periods'):
            if getattr(virt, _w, None) is None:
                node = source
                while node is not None:
                    _v = getattr(node, _w, None)
                    if _v is not None:
                        setattr(virt, _w, _v)
                        break
                    node = (getattr(node, '_parent', None)
                            or getattr(node, '_owner', None))
        # A short first period rides with the window it belongs to.
        from ...core.time import resolve_default_window
        _src_window = resolve_default_window(source)
        if _src_window is not None and _src_window.first_day is not None \
                and virt.__dict__.get('default_start') == _src_window.start:
            virt.__dict__['_default_first_day'] = _src_window.first_day
    # Bulk-assign components — the setter installs them into ``__dict__``
    # and populates ``_component_names`` in one shot. Direct per-key
    # writes (``virt._components[k] = v``) would write into the property
    # getter's snapshot, not the underlying storage.
    virt._components = {comp_name: comp for comp_name, comp in components}
    return virt


def _unique_sheet_names(roots: list[MultiVariableBase]) -> list[str]:
    """One Excel-safe (≤31 char), collision-free tab name per root, in order.

    ``sheet_name_for`` can return the same name for two sibling roots — a model
    and its ``.at(grain=…)`` clone both yield ``'Saas'`` (the projection copies
    ``display_name`` verbatim). openpyxl would de-duplicate the worksheet titles
    at ``create_sheet`` (``'Saas'`` → ``'Saas1'``), but that post-dedup name
    never propagated back into ``var_to_sheet`` / ``current_sheet``, so cross-tab
    references on the second tab lost their ``'Saas1'!`` prefix. Resolving the
    unique names here, once up front, keeps every layer in lockstep.

    On a collision, append the smallest integer suffix while TRUNCATING the base
    so ``base + suffix`` still fits the 31-char cap. The truncation is essential:
    a name already at the cap (``'X'*31``) would otherwise re-truncate
    ``base + '1'`` straight back to ``'X'*31`` and the loop would never escape.
    """
    seen: set[str] = set()
    names: list[str] = []
    for sheet_mv in roots:
        base = _safe_sheet_name(sheet_name_for(sheet_mv))
        name = base
        suffix = 1
        # Excel compares tab names case-insensitively: ``Sales`` and
        # ``SALES`` are one name, and openpyxl would rename the second
        # behind the formulas' back.
        while name.lower() in seen:
            tag = str(suffix)
            name = f"{base[:31 - len(tag)]}{tag}"
            suffix += 1
        seen.add(name.lower())
        names.append(name)
    return names


def _identity_homes(roots: "list[MultiVariableBase]",
                    addresses: "dict[str, VariableAddresses]"):
    """The emission tree's identity map (see ``identity_cells_for``).

    Homes are the rows whose cells genuinely HOLD the series: track
    subrows and untracked rows. The blend head of a tracked line is
    excluded — its cells are the splice, and a plan formula referencing
    a blend cell would move the moment an actual lands in the file.
    """
    from .blend_writer import TRACK_INFIX
    from .translator import identity_cells_for

    heads = {vid.split(TRACK_INFIX)[0] for vid in addresses
             if TRACK_INFIX in vid}
    rows = []
    for root in roots:
        for _name, comp in root.walk():
            vid = getattr(comp, 'id', None)
            if isinstance(comp, Variable) and vid and vid not in heads:
                rows.append((vid, comp))
    return identity_cells_for(rows, addresses)



def _build_var_to_sheet(
    addresses: dict[str, VariableAddresses],
    roots: list[MultiVariableBase],
    sheet_names: list[str] | None = None,
) -> dict[str, str]:
    """Map each variable's qualified-path id to the sheet name it lives on.

    The translator does case-sensitive equality on ``current_sheet`` vs
    ``var_to_sheet[ref]`` to decide whether to drop the sheet qualifier
    from a reference, so this MUST use the SAME post-dedup names the
    worksheets are created with. Pass ``sheet_names`` (from
    :func:`_unique_sheet_names`, aligned positionally with ``roots``) so
    two roots that derive the same :func:`sheet_name_for` (a model and its
    ``.at(grain=…)`` clone both yield ``'Saas'``) are distinguished
    ``'Saas'`` / ``'Saas1'`` — otherwise every genuinely cross-tab
    reference silently loses its prefix. When ``sheet_names`` is omitted
    the per-root :func:`sheet_name_for` value is used (back-compat for
    callers that don't pre-resolve unique names).
    """
    if sheet_names is None:
        sheet_names = [sheet_name_for(sheet_mv) for sheet_mv in roots]
    var_to_sheet: dict[str, str] = {}
    for sheet_mv, sheet_name in zip(roots, sheet_names):
        for _, comp in sheet_mv.walk():
            if isinstance(comp, Variable):
                var_id = _resolve_var_id(comp, addresses)
                if var_id:
                    var_to_sheet[var_id] = sheet_name
    return var_to_sheet


def _nesting_level(comp: Variable) -> int:
    """How deep under its TAB a variable sits: 0 for a first-depth
    row, 1 for a row inside a section, … Counted over the ownership
    chain to its top (the model root); the tab itself is the model's
    first depth, hence the ``- 2``."""
    hops = 0
    node = getattr(comp, '_owner', None)
    while node is not None:
        hops += 1
        node = getattr(node, '_parent', None)
    return max(0, hops - 2)


#: Number formats a period caption wears when it is the period's end
#: date (``=Time!C3``): what the text caption said, as a date format.
_LABEL_FORMATS = {
    'day': {'iso': 'yyyy-mm-dd', 'finance': 'dd mmm yyyy', 'compact': 'dd-mmm-yy'},
    'month': {'iso': 'yyyy-mm', 'finance': 'mmm yyyy', 'compact': 'mmm-yy'},
    'year': {'iso': 'yyyy', 'finance': 'yyyy', 'compact': 'yyyy'},
}


def _label_source(translator: Any, sheet_name: str, grain: str) -> Any:
    """``i → reference`` of the period-end cell a caption shows, when
    the workbook has a header carrying the period's end (this sheet's own
    row, else the anchor's); ``None`` keeps the text caption. A quarter
    has no date format to show "Q1 2026", so its caption stays text."""
    book = getattr(translator.renderer, 'time', None)
    if book is None or grain not in _LABEL_FORMATS:
        return None
    own = book.sheets.get(sheet_name)
    if own is None:
        return None
    vid = own.fields.get('end')
    if vid is None and own.anchor is not None and own.anchor.key == own.key:
        vid = own.anchor.fields.get('end')
    if vid is None:
        return None
    resolve = translator.renderer._resolve_var_addr
    return lambda i: resolve(vid, i, sheet_name)


def resolve_number_format(comp: Variable) -> str | None:
    """The cell number format a Variable's values display with — one
    resolution shared by the .xlsx writer and any other renderer that
    must display the same numbers the same way.

    Priority: the Variable's own ``_excel_layout.format`` → its own
    ``excel_props['number_format']`` → the nearest ancestor MV's
    ``excel_props['number_format']`` (a section states its unit
    convention once; a line of dates skips it unless it is a date
    format) → the format a coarser copy carries from its line, when it
    shows the number → a date format when the values are dates (with the
    clock when any has a time of day) → the datetime default → the
    resolved ExcelView's ``formats`` for numeric cells: by the line's
    unit, then the ``'number'`` floor. A ``formats['date']`` replaces the
    ISO date the values would otherwise call for.
    """
    layout = getattr(comp, "_excel_layout", None)
    if layout is not None:
        layout_format = getattr(layout, "format", None)
        if layout_format:
            return str(layout_format)
    props = getattr(comp, "_excel_props", None)
    if props and props.get("number_format"):
        return props["number_format"]
    date_format = _date_format_of(getattr(comp, "_value", None))
    node = getattr(comp, "_owner", None)
    while node is not None:
        props_up = getattr(node, "_excel_props", None)
        if props_up and props_up.get("number_format"):
            inherited = props_up["number_format"]
            # A section's format speaks for its numbers: a date line in
            # it keeps a date format unless the section's is one itself.
            if date_format is None or is_date_format(inherited):
                return inherited
            break
        node = getattr(node, "_parent", None)
    carried = getattr(comp, "_carried_number_format", None)
    if carried and (date_format is None or is_date_format(carried)) \
            and format_shows_a_number(carried):
        # A coarser copy (``line.at('year')``) wears its line's own format
        # where its sheet says none - unless the format hides the number
        # (a bar's mark over a count).
        return carried
    if date_format is not None or getattr(comp, "value_type", None) == "datetime":
        # A declared date look replaces the ISO default — unless the clock
        # shows, which a date format would hide.
        declared = (resolve_excel_view(comp).number_formats or {}).get("date")
        if declared and date_format != "yyyy-mm-dd hh:mm:ss":
            return declared
        return date_format or "yyyy-mm-dd"
    if getattr(comp, "value_type", None) in ("int", "float"):
        fmts = resolve_excel_view(comp).number_formats
        if fmts:
            # Unit-keyed first: ``formats={'$': …, '%': …}`` — declare
            # the unit on the line, the format follows it (and follows
            # DERIVED lines through unit algebra). Then ``'number'`` —
            # the generic floor for unitless numerics.
            unit = getattr(comp, "_unit", None)
            if unit is not None:
                by_unit = fmts.get(str(unit))
                if by_unit:
                    return by_unit
            number = fmts.get("number")
            if number:
                return number
        from ...core.time import time_field_of
        if time_field_of(comp) in ("days", "months", "index"):
            # A count of days or months spelled as date arithmetic
            # (end - start + 1) would otherwise show as a date.
            return "0"
    return None


def _date_format_of(value: object) -> str | None:
    """The date format a line's values call for, or ``None``.

    A date in a spreadsheet is a number that only LOOKS like a date
    through its format, so a line whose values are all dates (blanks
    aside) needs one — whether the dates were typed in or computed by a
    formula. The clock is shown when any value carries a time of day;
    otherwise 18:30 would silently read as the start of the day.
    """
    items = value if isinstance(value, list) else [value]
    present = [v for v in items if v is not None]
    if not present or not all(isinstance(v, date) for v in present):
        return None
    if any(isinstance(v, datetime) and v.time() != time() for v in present):
        return "yyyy-mm-dd hh:mm:ss"
    return "yyyy-mm-dd"


def _resolve_var_id(var: Variable, addresses: dict[str, VariableAddresses]) -> str | None:
    """Return the qualified-path id for ``var`` if it's in ``addresses``.

    Every Variable has a ``.path`` and layout writes addresses under
    ``str(var.path)``. Variables not in the dict (anonymous
    intermediates, out-of-scope references) return ``None``.
    """
    key = var.id
    return key if key in addresses else None


# ─── Cell-type colouring (``format_by_type``) ──────────────────────
#
# When an MV carries ``excel_props={'format_by_type': ...}``, every
# descendant value cell is tinted by what it IS — a hard input, a
# same-sheet formula, or a cross-sheet reference — the convention every
# professional financial model uses so a reviewer reads structure at a
# glance. The engine already knows each cell's type; this just colours it.

def _inherited_format_by_type(comp: Variable) -> object:
    """The active ``format_by_type`` setting for ``comp`` (or ``None``).

    A *truthy* resolved :class:`ExcelView` field (``True`` or a palette dict)
    wins; a falsy one (the house default's explicit ``False``, or ``None``)
    defers to the legacy ``excel_props['format_by_type']`` owner→parent walk, so
    models that set it the old way keep working. With neither set the result is
    ``None`` and cells are uncoloured.
    """
    view_setting = resolve_excel_view(comp).format_by_type
    if view_setting:
        return view_setting
    node = getattr(comp, "_owner", None)
    while node is not None:
        props = getattr(node, "_excel_props", None)
        if props and "format_by_type" in props:
            return props["format_by_type"]
        node = getattr(node, "_parent", None)
    return None


def _reaches_other_sheet(expr, sheet_name: str, var_to_sheet: dict, seen: set) -> bool:
    """True if ``expr`` references a placed cell on a different sheet.

    Recurses THROUGH floating intermediates (the anonymous Variables that
    operator chains like ``end > control.date`` produce) — those have no
    cell address of their own, so the real cross-sheet cell sits one or
    more levels below the formula's immediate refs.
    """
    for ref in expr.iter_refs():
        if id(ref) in seen:
            continue
        seen.add(id(ref))
        ref_sheet = var_to_sheet.get(ref.id)
        if ref_sheet is not None:
            if ref_sheet != sheet_name:
                return True
        elif ref._expr is not None:
            if _reaches_other_sheet(ref._expr, sheet_name, var_to_sheet, seen):
                return True
    return False


def _cell_type(comp: Variable, sheet_name: str, var_to_sheet: dict) -> str:
    """``input`` (literal), ``reference`` (formula reaching another sheet),
    or ``formula`` (same-sheet formula)."""
    if not comp.is_formula:
        return "input"
    if comp._expr is not None and _reaches_other_sheet(
        comp._expr, sheet_name, var_to_sheet, set()
    ):
        return "reference"
    return "formula"


def _type_color_for(comp: Variable, sheet_name: str, var_to_sheet: dict):
    """Font colour for a cell under an active ``format_by_type`` directive,
    or ``None`` when absent / the type is left uncoloured."""
    return type_style_for(comp, sheet_name, var_to_sheet).get("font_color")


def row_look(comp: Variable) -> dict:
    """A line's own look: the view's band for what the line is
    (``bands['input'|'link'|'formula']``, see :func:`row_kind`), then
    ``bands['result']`` when the model marks the line as its block's
    result (``excel_props={'result': True}``), then its own
    ``excel_props`` — what its label and its values wear. Public so every
    renderer dresses a line the same way."""
    own = getattr(comp, "_excel_props", None) or {}
    bands = resolve_excel_view(comp).band_styles or {}
    kind = row_kind(comp)
    look: dict = dict(bands.get(kind) or {})
    # A link never wears the result look: ``b.x = a.x.copy()`` carries the
    # source's ``excel_props``, mark included, and a line that opens a
    # calculation block by reading another block's result is not this
    # block's result.
    if own.get("result") and kind != "link":
        look.update(bands.get("result") or {})
    look.update(own)
    return look


def type_style_for(comp: Variable, sheet_name: str, var_to_sheet: dict) -> dict:
    """The look a value cell wears for what it IS — input, formula or
    cross-sheet reference — under an active ``format_by_type``
    directive, in the ``excel_props`` vocabulary; ``{}`` when none is
    active or the type is left bare. Public so every renderer dresses
    the same cell the same way.

    The palette's colour is the font colour; the view's
    ``cell_types['styles']`` adds the rest (a fill for inputs) and wins
    over the palette where both speak. The line's own ``excel_props``
    win over this at the call site.
    """
    config = _inherited_format_by_type(comp)
    if not config:
        return {}
    view = resolve_excel_view(comp)
    # ``format_by_type=True`` uses the resolved view's palette (seeded by the
    # house default ``mo.default_excel_view.type_colors``); a dict value is a
    # legacy inline palette and is used as-is.
    colors = config if isinstance(config, dict) else (view.type_colors or {})
    kind = _cell_type(comp, sheet_name, var_to_sheet)
    style: dict = {}
    color = colors.get(kind)
    if color:
        style["font_color"] = color
    style.update((view.type_styles or {}).get(kind) or {})
    return style


#: The font openpyxl gives a new book. A book that asks for it keeps
#: openpyxl's own Normal style — the file it always wrote.
_OPENPYXL_FONT = ("Calibri", 11.0)


def _font_props(base_font: Optional[dict]) -> dict:
    """A view's ``font`` group in the ``excel_props`` vocabulary."""
    out: dict = {}
    if base_font:
        if base_font.get("name"):
            out["font_family"] = base_font["name"]
        if base_font.get("size") is not None:
            out["font_size"] = base_font["size"]
    return out


def _set_book_font(wb: Workbook, font_props: dict) -> None:
    """Make ``font_props`` the book's Normal font — font 0, which every
    cell without a font of its own shows."""
    name = str(font_props.get("font_family") or _OPENPYXL_FONT[0])
    size = float(font_props.get("font_size") or _OPENPYXL_FONT[1])
    if (name, size) == _OPENPYXL_FONT:
        return
    font = Font(name=name, size=size)
    # openpyxl has no public door to the Normal font: font 0 of the book
    # and the Normal style's own font are the two places it lives.
    setattr(wb, "_fonts", IndexedList([font]))
    getattr(wb, "_named_styles")["Normal"].font = font


def _differs_from_book(font_props: dict, book_font: dict) -> dict:
    """The keys of a sheet's font that the book's Normal font does not
    already carry — what a cell of that sheet must say itself."""
    book = {
        "font_family": book_font.get("font_family") or _OPENPYXL_FONT[0],
        "font_size": float(book_font.get("font_size") or _OPENPYXL_FONT[1]),
    }
    out: dict = {}
    for key, value in font_props.items():
        mine = float(value) if key == "font_size" else value
        if book.get(key) != mine:
            out[key] = value
    return out


#: The ``excel_props`` keys that draw a row's LINE — its fill and its
#: borders — as opposed to its typography.
_LINE_KEYS = ("bg", "border_top", "border_bottom", "border_left", "border_right")
#: Everything of a look a cell wears: the line keys plus the typography
#: and the alignment (what ``_apply_styling`` reads).
_CELL_LOOK_KEYS = _LINE_KEYS + (
    "bold", "italic", "font_color", "font_size", "font_family", "text_align", "indent",
)


def lead_cells_for_row(
    comp: object,
    row: int,
    label_col: int,
    *,
    article: bool,
    unit: bool,
    fields: tuple = (),
    sum_resolver: object = None,
    number: Optional[str] = None,
) -> dict[tuple[int, int], str]:
    """The lead-column cells of ONE row: ``{(row, col): text}``.

    The article number in column A, the unit in its own column right
    after the label, then one column per projected field in declared
    order. Public so any other renderer places the same text in the
    same cells as the workbook — the geometry lives here once instead
    of being re-derived wherever it is needed, which is how two
    renderings drift apart.

    A ``('row_sum', caption)`` field is COMPUTED, not read: rows that
    opt in via ``excel_props={'row_sum': True}`` get whatever
    ``sum_resolver(comp)`` synthesizes — the writer passes a live
    ``=SUM(...)`` over the row's value cells, a value renderer the
    computed total — placed at the field's own column so every
    renderer agrees on geometry by construction.
    """
    out: dict[tuple[int, int], str] = {}
    if article:
        # The line's own article, else the number the view's numbering
        # gave it (``number``, from the layout's ``articles``).
        value = (getattr(comp, "_excel_props", None) or {}).get("article") or number
        if value:
            out[(row, 1)] = str(value)
    if unit:
        value = getattr(comp, "_unit", None)
        if value is not None:
            out[(row, label_col + 1)] = str(value)
    first = label_col + 1 + (1 if unit else 0)
    for i, (attr, _caption) in enumerate(fields or ()):
        if attr == "row_sum":
            flagged = has_row_sum(comp)
            text = (
                sum_resolver(comp)
                if flagged and callable(sum_resolver) else None
            )
            if text is not None and str(text) != "":
                out[(row, first + i)] = str(text)
            continue
        value = _read_field(comp, attr)
        if value is not None and str(value) != "":
            out[(row, first + i)] = str(value)
    return out


def has_row_sum(comp: object) -> bool:
    """Whether a line asks for its row total: its own ``excel_props``
    say so (either way), or the view's ``meta['row_sum_for']`` lists its
    re-grain rule. Public so every renderer totals the same lines."""
    own = getattr(comp, "_excel_props", None) or {}
    if "row_sum" in own:
        return bool(own["row_sum"])
    wanted = resolve_excel_view(comp).meta_row_sum_for
    if not wanted:
        return False
    recipe = getattr(getattr(comp, "_regrain", None), "default", None)
    return isinstance(recipe, str) and recipe in wanted


def format_shows_a_number(fmt: str) -> bool:
    """Whether a number format shows the number at all: its positive
    section holds a digit placeholder outside quoted text. A row in
    ``'"X";;'`` is a bar chart (a mark for any positive value, nothing
    else) and its total is a count the reader must still be able to
    read, so the total takes the row's format only when the format
    would show it."""
    section: list[str] = []
    quoted = False
    for ch in fmt:
        if ch == '"':
            quoted = not quoted
            continue
        if quoted:
            continue
        if ch == ";":
            break
        section.append(ch)
    text = "".join(section)
    return "General" in text or any(c in text for c in "0#?")


def format_shows_signed_numbers(fmt: "str | None") -> bool:
    """Whether a number format prints every SIGNED number as a number:
    its positive section holds a digit placeholder, and its negative
    one, when it has one, holds a digit placeholder and shows the sign
    (a minus, a parenthesis or a colour); the zero section may say
    anything (a dash is a line's own word for zero). ``'#,##0;;'`` hides negatives,
    ``'"X";;'`` draws a bar, ``';;;'`` hides everything and a date
    format reads a number as a date, so a difference of two values of
    such a line (the compare view's deviation column) is written in
    General instead. A renderer that shows the same column elsewhere
    should ask the same question of the same format, so every rendering
    prints the deviation alike."""
    if not fmt or fmt.strip().lower() == "general":
        return True
    if is_date_format(fmt):
        return False
    sections = _format_sections(fmt)
    return _section_shows_digits(sections[0]) and (
        len(sections) < 2
        or (_section_shows_digits(sections[1]) and _section_shows_sign(sections[1]))
    )


def _format_sections(fmt: str) -> "list[str]":
    """Quote-aware ``;`` split of a number format."""
    out: "list[str]" = []
    cur: "list[str]" = []
    quoted = False
    for ch in fmt:
        if ch == '"':
            quoted = not quoted
        if ch == ";" and not quoted:
            out.append("".join(cur))
            cur = []
            continue
        cur.append(ch)
    out.append("".join(cur))
    return out


def _section_shows_digits(section: str) -> bool:
    """A digit placeholder outside the literal machinery: quoted text,
    ``[...]`` tags, ``\\x`` escapes, ``_x`` pads and ``*x`` fills are
    literals, as Excel reads them."""
    return any(ch in "0#" for ch in _section_code(section))


_COLOUR_TAG = re.compile(
    r"\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)\]", re.IGNORECASE
)


def _section_shows_sign(section: str) -> bool:
    """A negative section that tells a negative apart: a minus or a
    parenthesis anywhere in it (quoted or escaped ones included), or a
    colour tag."""
    return "-" in section or "(" in section or bool(_COLOUR_TAG.search(section))


def _section_code(section: str) -> str:
    """The section with its literal machinery removed."""
    out: "list[str]" = []
    i, n = 0, len(section)
    while i < n:
        ch = section[i]
        if ch in ('"', "["):
            end = section.find('"' if ch == '"' else "]", i + 1)
            i = n if end == -1 else end + 1
            continue
        if ch in "_*\\":
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def row_sum_formula(refs: "list[str]") -> "str | None":
    """``=SUM(...)`` over a row's value cells, ranges compacted.

    The cells may be non-contiguous (subtotal columns interleave with
    the months), so consecutive same-row runs compress to ``B5:D5``
    pieces joined by commas — the formula a hand-built book would
    carry on the row's "Total" column.
    """
    if not refs:
        return None
    cells = sorted(coordinate_to_tuple(r) for r in refs)
    pieces: list[str] = []
    run_start = prev = None
    for r, c in cells:
        if prev is not None and r == prev[0] and c == prev[1] + 1:
            prev = (r, c)
            continue
        if run_start is not None:
            pieces.append(_range_piece(run_start, prev))
        run_start = prev = (r, c)
    if run_start is not None:
        pieces.append(_range_piece(run_start, prev))
    return f"=SUM({','.join(pieces)})"


def _range_piece(start: "tuple[int, int]", end: "tuple[int, int]") -> str:
    a = f"{get_column_letter(start[1])}{start[0]}"
    if start == end:
        return a
    return f"{a}:{get_column_letter(end[1])}{end[0]}"


def row_sum_eligible(comp: object, values_refs: "list[str]") -> bool:
    """Whether a flagged row honestly carries a total.

    Two refusals, public so every renderer applies the same ones: a
    SCALAR (single value cell) is not a run of periods — printing
    ``=SUM(D5)`` beside a number the row already shows is noise; an
    ALL-BLANK row would total a fabricated 0 where the honest answer
    is an empty cell.
    """
    if len(values_refs) < 2:
        return False
    return _has_any_number(getattr(comp, "value", None))


def _has_any_number(v: object) -> bool:
    if v is None:
        return False
    roles = getattr(v, "roles", None)
    if roles is not None:  # TrackValues — any track, any period
        return any(_has_any_number(v[r]) for r in roles)
    if isinstance(v, list):
        return any(
            isinstance(x, (int, float)) and not isinstance(x, bool)
            for x in v
        )
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _read_field(comp: object, attr: str) -> object | None:
    """One projected field off a row, by name.

    Looks where a value of that name can honestly live: the public
    attribute, its private twin, then ``excel_props``. Anything not
    found leaves the cell empty — a projected column reports what the
    model holds and never invents a value for it.
    """
    for name in (attr, f"_{attr}"):
        value = getattr(comp, name, None)
        if value is not None and not callable(value):
            return value
    props = getattr(comp, "_excel_props", None) or {}
    return props.get(attr)


def _write_captions(
    ws: Worksheet,
    row: int,
    props: dict,
    meta: tuple,
    meta_fields: tuple,
    constants_caption: object,
    const_col: int,
    opening_caption: object,
    open_col: int | None,
) -> None:
    """The column captions of a sheet's header row.

    Lead metadata columns get their own headers — a column nobody
    labelled is a column the reader has to guess at, and the period
    labels deliberately start AFTER them. The flag IS the caption:
    ``meta={'unit': 'Unit'}`` heads the column with that text, while
    a bare ``True`` keeps the historical headerless column. Written only
    where the layout gave the sheet a header row (a period header, or the
    caption row of a sheet of constants — ``caption_rows_by_sheet``):
    elsewhere the layout puts DATA in row 1, and a caption there would
    land on top of it.
    """
    article, unit, label = meta
    if isinstance(article, str) and article:
        _apply_styling(ws.cell(row=row, column=1, value=article), props)
    if isinstance(label, str) and label:
        _apply_styling(
            ws.cell(row=row, column=1 + (1 if article else 0), value=label), props,
        )
    # The constants column's header — "Value" in the book form. Sits
    # immediately left of the first period, where an unlabeled run of
    # scalars would look like the sheet had slipped against its own
    # timeline.
    if isinstance(constants_caption, str) and constants_caption and const_col >= 1:
        _apply_styling(ws.cell(row=row, column=const_col, value=constants_caption), props)
    # The opening column's header — "Pre-start" — just left of the first
    # period.
    if isinstance(opening_caption, str) and opening_caption and open_col:
        _apply_styling(ws.cell(row=row, column=open_col, value=opening_caption), props)
    first = 2 + (1 if article else 0) + (1 if unit else 0)
    for i, (_attr, caption) in enumerate(meta_fields):
        if isinstance(caption, str) and caption:
            _apply_styling(ws.cell(row=row, column=first + i, value=caption), props)
    if isinstance(unit, str) and unit:
        _apply_styling(
            ws.cell(row=row, column=2 + (1 if article else 0), value=unit), props,
        )


def _write_sheet(
    ws: Worksheet,
    sheet_mv: MultiVariableBase,
    addresses: dict[str, VariableAddresses],
    translator: ExcelTranslator,
    sheet_name: str,
    section_header_rows: dict[str, tuple],
    time_headers: dict[int, tuple],
    engine_meta: dict | None = None,
    board_var_ids: set | None = None,
    section_header_echoes: dict[str, list] | None = None,
    inline_section_ids: set | None = None,
    period_start: int = 0,
    period_var_ids: set | None = None,
    meta_fields: tuple = (),
    constants_caption: object = None,
    totals_plan: object = None,
    group_plan: tuple | None = None,
    row_outline: dict | None = None,
    header: Any = None,
    book_font: dict | None = None,
    columns: dict | None = None,
    top: int = 0,
    opening_caption: object = None,
    freeze: str | None = None,
    articles: dict | None = None,
    caption_row: int | None = None,
) -> None:
    """Fill a worksheet: an optional timeline header row, section-header
    labels for nested MVs, then a row per Variable (label + per-period
    values or formulas)."""
    _sheet_view = resolve_excel_view(sheet_mv)
    # What the sheet's cells must say about their font themselves: the
    # keys of the sheet's font the book's Normal font does not carry.
    # Nothing, on a sheet that uses the book's font — the usual case.
    _sheet_font = _differs_from_book(
        _font_props(_sheet_view.base_font), book_font or {}
    )
    _full_rows = bool(_sheet_view.sheet_full_rows)
    # Each lead column's own look, by column number (``sheet.columns``).
    _col_style: dict[int, dict] = {}
    for _role, _spec in (_sheet_view.sheet_columns or ()):
        _c = (columns or {}).get(_role)
        _look = {k: v for k, v in _spec.items() if k != "width"}
        # Lead columns only: the label, constants, opening and period
        # columns hold the row's own cells, which wear the row's look.
        if _c and _look and _role not in ("label", "constants", "opening", "period"):
            _col_style[_c] = _look
    # Section headers first — every nested MV with a recorded slot
    # gets its display_name at (row, col). In flatten mode, nested
    # ``_is_sheet=True`` MVs render as sections inside their parent's
    # sheet, so the ``_is_sheet`` flag isn't a filter here.
    #
    # MV-level ``excel_props`` styling (``bold`` / ``bg`` / ``font_color``
    # / ``border_*`` / etc.) lands on the header cell — same vocabulary
    # the per-Variable path uses, so authoring a styled section title
    # is just ``mo.MultiVariable("ASSETS", excel_props={'bold': True,
    # 'bg': '#0D1F3C', 'font_color': '#FFFFFF'})``. Replaces the
    # workaround of declaring a phantom ``mo.Variable(0.0, ...)``
    # row purely to get a styled label.
    # Rightmost data column on this sheet — section-header styling extends
    # across the whole row to that column so a styled section reads as a
    # full-width band (not just a coloured label cell).
    sheet_max_col = 1
    for addr in addresses.values():
        for cell_addr in addr.values:
            _, col_idx = coordinate_to_tuple(cell_addr)
            if col_idx > sheet_max_col:
                sheet_max_col = col_idx
    # THIS sheet's own last column (``sheet_max_col`` spans the widest
    # sheet of the book): where a full-width row stops.
    _own_max_col = max((columns or {}).values(), default=1)
    _own_ids = [
        _resolve_var_id(c, addresses) for _, c in sheet_mv.walk()
        if isinstance(c, Variable)
    ] + [m.id for _s, m in (header.mirrors if header is not None else ())]
    for _vid in _own_ids:
        if _vid in addresses:
            for cell_addr in addresses[_vid].values:
                _own_max_col = max(_own_max_col, coordinate_to_tuple(cell_addr)[1])
    # The totals interleaved with the periods are this sheet's columns
    # too, the closing ones included: a final quarter's and year's total
    # stand right of every line's own value cell, and a full-width row's
    # line and a section band stopped one column short of them.
    _plan_cols = getattr(totals_plan, "col_of", None)
    _totals_end = (
        max(_plan_cols.values()) + max(getattr(totals_plan, "stride", 1), 1) - 1
        if _plan_cols else 0
    )
    _own_max_col = max(_own_max_col, _totals_end)

    # Identity lookup first — worksheet titles are sanitized/uniquified
    # AFTER layout, so a name key can silently miss ('P/L' → 'P_L').
    _meta_triple = (engine_meta or {}).get(
        id(sheet_mv), (engine_meta or {}).get(sheet_name, (False, False, False))
    )
    # Tolerant unpack: the layout has carried a 2-tuple historically.
    _meta_article, _meta_unit = _meta_triple[0], _meta_triple[1]
    _meta_label = _meta_triple[2] if len(_meta_triple) > 2 else False

    # Timeline header — period labels across row 1 (cols B onward), one per
    # value column on THIS sheet. Present only when the layout resolved a time
    # window for the sheet (recorded in ``time_headers``); a window-less sheet
    # keeps row 1 for data and renders byte-identically to before.
    header_info = time_headers.get(id(sheet_mv))
    if header_info is not None:
        _hdr_first = period_start or (
            2 + (1 if _meta_article else 0) + (1 if _meta_unit else 0)
        )
        from ...core.time import _period_labels
        start, grain, style = header_info
        # This sheet's own rightmost data column — the header must not paint
        # labels past real values (the global ``sheet_max_col`` above can span
        # a wider neighbour sheet).
        header_max_col = 1
        for _, comp in sheet_mv.walk():
            if isinstance(comp, Variable):
                vid = _resolve_var_id(comp, addresses)
                if vid is not None and vid in addresses:
                    if board_var_ids and vid in board_var_ids:
                        # Board columns are INSTANCES, not periods —
                        # the timeline must not stretch over them.
                        continue
                    if period_var_ids is not None and vid not in period_var_ids:
                        # Nor does a CONSTANT belong under a period
                        # label: it holds one number that is not a
                        # January figure, and the layout parked it in
                        # its own column left of the periods.
                        continue
                    for cell_addr in addresses[vid].values:
                        _, col_idx = coordinate_to_tuple(cell_addr)
                        header_max_col = max(header_max_col, col_idx)
        for _src, _mirror in (header.mirrors if header is not None else ()):
            for cell_addr in addresses[_mirror.id].values:
                header_max_col = max(header_max_col, coordinate_to_tuple(cell_addr)[1])
        n_cols = header_max_col - _hdr_first + 1 if _hdr_first else 0
        if n_cols > 0:
            # Header band: the resolved view's ``bands['header']`` styles
            # the period-label row (brand fill, font colour) over the
            # bold floor — the book's masthead, declared once.
            _hdr_bands = (_sheet_view.band_styles or {})
            _hdr_props = {**_sheet_font, 'bold': True, **_hdr_bands.get('header', {})}
            _hdr_start = _hdr_first
            if totals_plan is not None:
                # HIERARCHICAL header — "year" over "quarters" over
                # "months": each bucket is a merged span over its
                # months, its finer buckets, and its own total column
                # (which therefore carries no month-row label).
                from .totals import header_rows

                _plan = totals_plan
                _month_lbls = _period_labels(
                    start, None, _plan.n, grain, style
                )
                # NATIVE Excel column groups: a month column's
                # outline level = how many buckets enclose it (year+
                # quarter → 2), a quarter's total column sits one
                # level up (inside the year only), the year's total
                # column stays ungrouped. Excel then draws its own
                # ± over the columns, and collapsing a year works in
                # the workbook itself. Summary RIGHT of detail — where
                # our totals stand.
                ws.sheet_properties.outlinePr.summaryRight = True
                _kinds_depth = {"quarter": 0, "year": 0}
                _enclosers = [k for k in ("year", "quarter")
                              if k in _plan.buckets]
                # A period is a GROUP of ``stride`` columns and the
                # bracket must cover all of them — a level on the
                # first column alone drew Excel brackets over one
                # column in four, and the grouping read as noise.
                # Hidden operand columns go one level deeper: their
                # own collapsed subgroup, so expanding a quarter does
                # not unhide what the track selection deliberately hid.
                _stride_o = group_plan[0] if group_plan else 1
                _hidden_o = set(group_plan[2]) if group_plan else set()
                for _ci, (_kind, _bi) in enumerate(_plan.columns):
                    _col_num = _plan.col_of[(_kind, _bi)]
                    if _kind == "month":
                        _level = len(_enclosers)
                    elif _kind == "quarter":
                        _level = 1 if "year" in _plan.buckets else 0
                    else:
                        _level = 0
                    for _si in range(_stride_o):
                        _lv = _level + (1 if _si in _hidden_o else 0)
                        if _lv > 0:
                            ws.column_dimensions[
                                get_column_letter(_col_num + _si)
                            ].outline_level = min(7, _lv)
                for hrow in header_rows(_plan, _month_lbls):
                    for cs in hrow["cells"]:
                        # Tier spans read from their LEFT edge — a
                        # centered "2026" floats far from the row
                        # names it belongs to. The month row stays
                        # centered under it.
                        _align = (
                            'center' if cs["colspan"] == 1 else 'left'
                        )
                        _apply_styling(
                            ws.cell(
                                row=hrow["row"] + top,
                                column=cs["col"],
                                value=cs["label"],
                            ),
                            {**_hdr_props, 'bold': True,
                             'text_align': _align},
                        )
                        _rspan = cs.get("rowspan", 1)
                        if cs["colspan"] > 1 or _rspan > 1:
                            ws.merge_cells(
                                start_row=hrow["row"] + top,
                                start_column=cs["col"],
                                end_row=hrow["row"] + top + _rspan - 1,
                                end_column=cs["col"] + cs["colspan"] - 1,
                            )
            else:
                _stride = group_plan[0] if group_plan else 1
                _n_periods = max(1, n_cols // _stride)
                _end_ref = _label_source(translator, sheet_name, grain)
                for i, lbl in enumerate(
                    _period_labels(start, None, _n_periods, grain, style)
                ):
                    _col0 = _hdr_start + i * _stride
                    _label_cell = ws.cell(row=1 + top, column=_col0, value=lbl)
                    if _end_ref is not None:
                        # The caption IS the period's end date, formatted:
                        # move the anchor's start and it follows.
                        _label_cell.value = f"={_end_ref(i)}"
                        _label_cell.number_format = _LABEL_FORMATS[grain][style]
                    _apply_styling(_label_cell, _hdr_props)
                    if _stride > 1:
                        ws.merge_cells(
                            start_row=1 + top, start_column=_col0,
                            end_row=1 + top, end_column=_col0 + _stride - 1,
                        )
            # Track column groups: the word row under the period
            # labels — "plan | actual/forecast | …", the track labels over
            # every group. Unselected-but-needed columns (a blend
            # operand the splice formula references)
            # hide themselves rather than widen the visible book.
            if group_plan is not None:
                _g_stride, _g_words, _g_hidden = group_plan
                _word_row = top + (
                    totals_plan.depth + 1 if totals_plan is not None else 2
                )
                _group_starts = (
                    [totals_plan.col_of[k] for k in totals_plan.columns]
                    if totals_plan is not None
                    else [
                        _hdr_start + g * _g_stride
                        for g in range(max(1, n_cols // _g_stride))
                    ]
                )
                for _gc in _group_starts:
                    for _si, _w in enumerate(_g_words):
                        if _w:
                            _apply_styling(
                                ws.cell(row=_word_row, column=_gc + _si,
                                        value=_w),
                                {**_hdr_props, 'font_size': 9},
                            )
                    for _si in _g_hidden:
                        ws.column_dimensions[
                            get_column_letter(_gc + _si)
                        ].hidden = True

            # On a hierarchical header the captions belong to the
            # BOTTOM row — beside the month names, not the year.
            _write_captions(
                ws, top + (totals_plan.depth if totals_plan is not None else 1),
                _hdr_props, (_meta_article, _meta_unit, _meta_label), meta_fields,
                constants_caption,
                (columns or {}).get('constants') or _hdr_first - 1,
                opening_caption, (columns or {}).get('opening'),
            )
    elif caption_row is not None:
        # A sheet of constants in a book whose sheets share their
        # columns: no period header, the same captions all the same.
        _write_captions(
            ws, caption_row,
            {**_sheet_font, 'bold': True,
             **(_sheet_view.band_styles or {}).get('header', {})},
            (_meta_article, _meta_unit, _meta_label), meta_fields,
            constants_caption, (columns or {}).get('constants') or 0,
            None, None,
        )

    _hidden_rows: set[int] = set()
    for _, comp in sheet_mv.walk():
        if isinstance(comp, MultiVariableBase):
            section_id = comp.id or comp.python_name or comp._name_in_parent
            pos = section_header_rows.get(section_id)
            if pos is None or not comp.display_name:
                continue
            if effective_hidden(comp):
                _hidden_rows.add(pos[0])
            # A section can legitimately name itself in more than one
            # place — a record laid out as a table row also labels its
            # row in each per-period block below. Same label, same
            # styling, one pass.
            slots = [pos, *((section_header_echoes or {}).get(section_id, []))]
            # Dense nesting shows hierarchy by INDENT alone (its gap
            # rows are zero) — and only variable labels were indented,
            # so a sub-section's title sat flush left and read as a
            # sibling of its own parent. Same rule, same amount: the
            # view's indent × the section's depth under the tab.
            _hdr_label = comp.display_name
            _hdr_indent = resolve_excel_view(comp).nesting_indent
            # MVs chain by ``_parent`` (``_owner`` is a Variable's
            # link) — same ``hops - 2`` as ``_nesting_level``: the
            # model root and the tab are structure, not depth.
            _hops = 0
            _node = getattr(comp, '_parent', None)
            while _node is not None:
                _hops += 1
                _node = getattr(_node, '_parent', None)
            _hdr_level = max(0, _hops - 2)
            if _hdr_indent and _hdr_level > 0:
                _hdr_label = " " * (_hdr_indent * _hdr_level) + _hdr_label
            for row, col in slots:
                header_cell = ws.cell(
                    row=row, column=col, value=_hdr_label
                )
                # Band styling: the resolved view's ``band_styles[role]`` provides
                # defaults UNDER the MV's own excel_props (explicit props win, so
                # a view-less model is byte-identical). Role derives from DEPTH:
                # first level under the tab is a ``section``, deeper a
                # ``subsection``. Depth, not an own ``bg``, decides the role, so
                # ``bands.section.bg`` can paint a section that has no fill of
                # its own.
                own = comp._excel_props or {}
                if _meta_article and own.get('article'):
                    ws.cell(row=row, column=1, value=str(own['article']))
                role = "section" if _hdr_level == 0 else "subsection"
                view_bands = resolve_excel_view(comp).band_styles or {}
                band_props = {**_sheet_font, **view_bands.get(role, {}), **own}
                if band_props:
                    _apply_styling(header_cell, band_props)
                    # Extend the band across the row (label cell already
                    # styled) — unless the label SHARES its row with
                    # data, as a record's does in a records-oriented
                    # table: there the banner would paint the numbers.
                    # A full-width row starts at the first column and
                    # stops at this sheet's own last one.
                    if section_id not in (inline_section_ids or set()):
                        _band_cols = (
                            range(1, _own_max_col + 1) if _full_rows
                            else range(col + 1, max(sheet_max_col, _totals_end) + 1)
                        )
                        for c in _band_cols:
                            if c != col:
                                _apply_styling(ws.cell(row=row, column=c), band_props)

    # The title row: the sheet's name over its label column, its
    # description beside it — where the rows' descriptions stand.
    if top:
        _bands = _sheet_view.band_styles or {}
        for (_r, _c), (_text, _role) in title_cells(
            sheet_mv, _sheet_view, columns or {}, sheet_name,
        ).items():
            _default = {'bold': True} if _role == 'title' else {'italic': True}
            _apply_styling(ws.cell(row=_r, column=_c, value=_text),
                           {**_sheet_font, **(_bands.get(_role) or _default)})

    # Header mirrors — the anchor's rows at the top of this sheet, in
    # the view's ``bands['mirror']`` look (italic when none is declared).
    _mirror_look = {
        **_sheet_font,
        **((_sheet_view.band_styles or {}).get('mirror') or {'italic': True}),
    }
    for _src, _mirror in (header.mirrors if header is not None else ()):
        _maddr = addresses[_mirror.id]
        if _maddr.name:
            ws[_maddr.name] = _src.display_name
            _apply_styling(ws[_maddr.name], _mirror_look)
            if _meta_unit:
                # A header row says what it counts, like any other row.
                _mrc = coordinate_to_tuple(_maddr.name)
                for (_r, _c), _text in lead_cells_for_row(
                    _mirror, _mrc[0], _mrc[1], article=False, unit=True,
                ).items():
                    _apply_styling(ws.cell(row=_r, column=_c, value=_text),
                                   {**_col_style.get(_c, {}), **_mirror_look})
        # A mirror shows its row's numbers, so it shows them the way the
        # row does — the row's own tree decides the format, not the
        # detached mirror's.
        _mfmt = resolve_number_format(_src)
        # On a sheet of slot groups a mirror is one cell of each period,
        # at the group's first slot: its band runs on across the group,
        # as a line's does (below).
        _mline = (
            {k: _mirror_look[k] for k in _LINE_KEYS if k in _mirror_look}
            if group_plan is not None and group_plan[0] > 1 and not _full_rows
            else {}
        )
        for i, cell_addr in enumerate(_maddr.values):
            cell = ws[cell_addr]
            cell.value = _cell_content(_mirror, i, _mirror.id, translator, sheet_name, _maddr)
            if _mfmt:
                cell.number_format = _mfmt
            _apply_styling(cell, _mirror_look)
            if _mline:
                _mr, _mc = coordinate_to_tuple(cell_addr)
                for _c in range(_mc + 1, _mc + group_plan[0]):
                    _apply_styling(ws.cell(row=_mr, column=_c), _mline)
        _mopen = opening_content(_mirror, translator, sheet_name)
        if _mopen is not None and _maddr.opening:
            cell = ws[_maddr.opening]
            cell.value = _mopen
            if _mfmt:
                cell.number_format = _mfmt
            _apply_styling(cell, _mirror_look)

    # Variable cells.
    for _, comp in sheet_mv.walk():
        if not isinstance(comp, Variable):
            continue

        var_id = _resolve_var_id(comp, addresses)
        if var_id is None:
            continue
        addr = addresses[var_id]
        if effective_hidden(comp):
            for _cell in ([addr.name] if addr.name else []) + list(addr.values):
                _hidden_rows.add(coordinate_to_tuple(_cell)[0])

        label = comp.display_name
        # Lead metadata columns: the article number in column A, the
        # unit in its own column right after the label. Written below,
        # once the row's look is known.
        _lead: dict = {}
        if addr.name:
            _lrc = coordinate_to_tuple(addr.name)
            _lead = lead_cells_for_row(
                comp,
                _lrc[0],
                _lrc[1],
                article=bool(_meta_article),
                unit=bool(_meta_unit),
                fields=meta_fields,
                # The book's row total is a LIVE formula over the row's
                # own value cells — openpyxl stores an ``=``-string as
                # a formula, so the .xlsx recalculates it on edit.
                # Eligibility refuses scalars and all-blank rows (see
                # ``row_sum_eligible``).
                sum_resolver=lambda _c_, _a=addr: (
                    row_sum_formula(list(_a.values))
                    if row_sum_eligible(_c_, list(_a.values)) else None
                ),
                number=(articles or {}).get(var_id),
            )
        # Dense-nesting indent: the view's ``nesting.indent`` spaces
        # per level under the tab — hierarchy read from the indent,
        # not from whitespace rows.
        _indent = resolve_excel_view(comp).nesting_indent
        if _indent and label:
            _level = _nesting_level(comp)
            if _level > 0:
                label = " " * (_indent * _level) + label
        # Suffix the unit when the Variable has one — "Revenue ($)".
        # (Skipped when the unit renders as its own lead column.)
        unit = None if _meta_unit else getattr(comp, '_unit', None)
        if unit is not None:
            unit_str = str(unit)
            if unit_str:
                label = f"{label} ({unit_str})"
        if label and addr.name:
            ws[addr.name] = label

        # Optional per-variable number format (e.g. dates → "mmm yyyy").
        # Two override sources, checked in priority order:
        #   1. ``_excel_layout.format`` — a structured layout tag an
        #      extension may attach. Wins because it's the explicit
        #      per-Variable format on the layout system.
        #   2. ``_excel_props['number_format']`` — legacy flag-dict path.
        # Engine reads the attribute opaquely (duck-typed: any object
        # with a ``format`` attribute); the layout type is not defined here.
        number_format = resolve_number_format(comp)

        props = row_look(comp)
        # Default font from the resolved ExcelView (declared once on an ancestor
        # and cascaded). Injected UNDER the Variable's own props so an explicit
        # font_family / font_size always wins; absent (no view) → today's default
        # so output stays byte-identical.
        _base_font = resolve_excel_view(comp).base_font
        if _base_font:
            _font_defaults: dict = {}
            if _base_font.get("name"):
                _font_defaults["font_family"] = _base_font["name"]
            if _base_font.get("size") is not None:
                _font_defaults["font_size"] = _base_font["size"]
            if _font_defaults:
                props = {**_font_defaults, **props}
        # Cell-type colouring (``format_by_type`` on an ancestor MV): tint the
        # VALUE cells by type — input / formula / cross-sheet reference — so a
        # reviewer reads the model's structure at a glance. The label keeps its
        # own style; an explicit ``font_color`` on the Variable always wins.
        value_props = props
        type_style = type_style_for(comp, sheet_name, translator.var_to_sheet)
        if type_style:
            value_props = {**type_style, **props}

        # A sheet of slot groups (a plan-vs-actual book: actual | plan | var per
        # period) draws a line as ONE cell of each period: its rule and
        # fill run across the whole group. The slots its tracks and
        # metrics fill wear the line themselves; the empty ones — every
        # slot but the first on a line without tracks, a metric with
        # nothing to say — take the line keys here. Laid before the
        # line's own cells; a full-width sheet runs the line everywhere
        # below.
        if (
            group_plan is not None and group_plan[0] > 1 and not _full_rows
            and getattr(comp, '_col_head', None) is None
            and (period_var_ids is None or var_id in period_var_ids)
            and addr.values
        ):
            _gline = {k: props[k] for k in _LINE_KEYS if k in props}
            if _gline:
                _slot0 = getattr(comp, '_col_slot', 0) or 0
                for cell_addr in addr.values:
                    _gr, _gc = coordinate_to_tuple(cell_addr)
                    _g0 = _gc - _slot0
                    # The group starts the slot's offset back — where the
                    # layout put the cell at its slot. A one-period line
                    # without totals sits at the first period column
                    # instead; there is no group to draw then.
                    if _g0 < max(period_start or 0, 1):
                        continue
                    for _c in range(_g0, _g0 + group_plan[0]):
                        if _c != _gc:
                            _apply_styling(ws.cell(row=_gr, column=_c), _gline)
        for i, cell_addr in enumerate(addr.values):
            cell = ws[cell_addr]
            cell.value = _cell_content(comp, i, var_id, translator, sheet_name, addr)
            if number_format:
                cell.number_format = number_format
            _apply_styling(cell, value_props)
        # The row's opening cell — filled by the time rows only.
        if addr.opening:
            _open = opening_content(comp, translator, sheet_name)
            if _open is not None:
                cell = ws[addr.opening]
                cell.value = _open
                if number_format:
                    cell.number_format = number_format
                _apply_styling(cell, value_props)
        # Value-driven fill (``bg_nonzero``): a LIVE conditional rule —
        # nonzero NUMBERS wear the color, zeros, blanks and text stay
        # bare (a CellIs "notEqual 0" rule would also paint error
        # tokens and any text) — so the band repaints itself
        # when the file's values are edited. A multi-range sqref
        # because totals columns interleave with the months; the
        # relative anchor is the range's first cell and Excel walks
        # it across every covered cell.
        _bg_nz = _normalize_color(props.get("bg_nonzero"))
        if _bg_nz and addr.values:
            from openpyxl.formatting.rule import FormulaRule
            _anchor = min(addr.values, key=coordinate_to_tuple)
            ws.conditional_formatting.add(
                " ".join(addr.values),
                FormulaRule(
                    formula=[
                        f"AND(ISNUMBER({_anchor}),{_anchor}<>0)"
                    ],
                    fill=PatternFill(
                        start_color=_bg_nz, end_color=_bg_nz,
                        fill_type="solid",
                    ),
                ),
            )
        # Apply the same styling to the label cell so the row reads as
        # one styled unit — when a user marks a Variable bold, both the
        # label and its values bold (matches Excel users' expectations
        # that styling describes the "row" not individual cells).
        if label and props and addr.name:
            _apply_styling(ws[addr.name], props)
        # The lead cells: each in its column's own look. The row total is
        # a number of the row — it shows in the row's number format, and
        # on a full-width row it wears the row's typography too; there the
        # line's fill and borders also run across every other cell.
        _sum_col = (columns or {}).get('row_sum')
        _line = (
            {k: props[k] for k in _LINE_KEYS if k in props} if _full_rows else {}
        )
        for (_r, _c), _text in _lead.items():
            _cell = ws.cell(row=_r, column=_c, value=_text)
            _look = {**_sheet_font, **_col_style.get(_c, {})}
            if _c == _sum_col:
                if number_format and format_shows_a_number(number_format):
                    _cell.number_format = number_format
                if _full_rows:
                    _look.update(props)
            elif _line:
                _look.update(_line)
            if _look:
                _apply_styling(_cell, _look)
        if _line and addr.name:
            _row = coordinate_to_tuple(addr.name)[0]
            _taken = {c for (_r, c) in _lead} | {
                coordinate_to_tuple(a)[1] for a in [addr.name, *addr.values]
            }
            for _c in range(1, _own_max_col + 1):
                if _c not in _taken:
                    _apply_styling(ws.cell(row=_row, column=_c), _line)

    _auto_width_first_column(
        ws,
        label_col=1 + (1 if _meta_article else 0),
        meta_article=_meta_article,
        meta_unit=_meta_unit,
    )
    # Declared widths, by column role, over the automatic ones.
    for _c, _w in column_widths(_sheet_view, columns or {}, _own_max_col).items():
        ws.column_dimensions[get_column_letter(_c)].width = _w
    if _sheet_view.sheet_gridlines is False:
        ws.sheet_view.showGridLines = False
    # Freeze the label column so row labels stay visible while scrolling across
    # many period columns — standard for wide financial models. When a timeline
    # header occupies row 1, freeze it too (``B2``) so the period labels stay
    # pinned alongside the row labels.
    _freeze_col = get_column_letter(
        period_start
        or 2 + (1 if _meta_article else 0) + (1 if _meta_unit else 0)
    )
    _hdr_depth = (
        getattr(totals_plan, "depth", 1) if totals_plan is not None else 1
    ) + (1 if group_plan is not None else 0)
    _freeze_row = 1 + top + (_hdr_depth if id(sheet_mv) in time_headers else 0)
    if header is not None and header.mirrors and id(sheet_mv) in time_headers:
        # The header rows are part of the masthead: freeze below them.
        _freeze_row = max(
            coordinate_to_tuple(addresses[m.id].values[0])[0]
            for _s, m in header.mirrors
        ) + 1
    ws.freeze_panes = freeze or f"{_freeze_col}{_freeze_row}"
    # Native ROW groups following the section tree — a collapsible
    # outline in the file itself. The section header is the group's
    # summary and sits ABOVE it, so the ± lands on the header row
    # (summaryBelow=False), beside the label it collapses.
    if row_outline:
        ws.sheet_properties.outlinePr.summaryBelow = False
        for _r, _lv in row_outline.items():
            ws.row_dimensions[_r].outline_level = min(7, _lv)
    # Helper rows (``excel_props={'hidden': True}``, inherited down
    # the tree): written in full — values, formulas, labels — but
    # natively hidden. The book stays auditable; the eye stays clear.
    for _r in _hidden_rows:
        ws.row_dimensions[_r].hidden = True


# ─── Per-cell styling (``excel_props`` flag dict) ──────────────
#
# Engine consumes a small fixed vocabulary off ``_excel_props``:
# typography (``bold``, ``italic``, ``font_color``, ``font_size``,
# ``font_family``), fill (``bg``), borders (``border_top`` /
# ``_bottom`` / ``_left`` / ``_right``), alignment (``text_align``,
# ``indent``). Unknown keys (validated by ``Component._EXCEL_PROP_KEYS``
# at construction time) never reach here. Number format is handled
# upstream in ``_write_sheet``; this helper is for the visual style.


# Border thickness vocabulary. Truthy ``bool`` defaults to ``"thin"`` —
# matches Excel's default border weight when a user clicks the border
# button without picking a style.
_BORDER_STYLES = {
    "thin",
    "medium",
    "thick",
    "double",
    "dotted",
    "dashed",
    "hair",
    "mediumDashed",
}


def _normalize_color(value: object) -> str | None:
    """Coerce a color value to openpyxl's ``'RRGGBB'`` / ``'AARRGGBB'``
    8-char hex form. Accepts ``'#RRGGBB'`` / ``'RRGGBB'`` /
    ``'#RRGGBBAA'`` / ``'RRGGBBAA'`` and strips the leading ``#``.

    Returns ``None`` for unrecognised shapes so the caller skips
    applying the style instead of crashing on bad input.
    """
    if not isinstance(value, str):
        return None
    v = value.strip().lstrip("#").upper()
    if len(v) == 6:
        # Prepend a fully-opaque alpha. openpyxl otherwise pads a 6-char
        # RGB with ``00`` alpha (``00RRGGBB``), which some viewers render
        # transparent; ``FF`` matches how Excel stores hand-set fills.
        return "FF" + v
    if len(v) == 8:
        # Conservative — could regex \\A[0-9A-F]+\\Z but openpyxl
        # raises a clear error on bad chars too. Trust the input.
        return v
    return None


def _border_side(value: object) -> Side | None:
    """Map a border-prop value to an openpyxl :class:`Side` (or None
    when no border should render). ``True`` → thin; explicit named
    style → that style; falsy / unrecognised → no border."""
    if value is True:
        return Side(style="thin")
    if isinstance(value, str) and value in _BORDER_STYLES:
        return Side(style=value)
    return None


def _apply_styling(cell: object, props: dict) -> None:
    """Apply typography / fill / border / alignment from ``props``.

    Each key is independently optional; missing keys leave the cell's
    default style alone. Existing styles (e.g. ``number_format`` set
    earlier) are preserved — openpyxl style objects are immutable
    composites, so we build a new Font/Fill/etc. each time rather
    than mutating in place.
    """
    if not props:
        return

    # Typography.
    font_kwargs: dict = {}
    if props.get("bold"):
        font_kwargs["bold"] = True
    if props.get("italic"):
        font_kwargs["italic"] = True
    fc = _normalize_color(props.get("font_color"))
    if fc:
        font_kwargs["color"] = fc
    fs = props.get("font_size")
    if isinstance(fs, (int, float)):
        font_kwargs["size"] = float(fs)
    ff = props.get("font_family")
    if isinstance(ff, str) and ff:
        font_kwargs["name"] = ff
    if font_kwargs:
        # Merge over the cell's existing Font so we don't drop
        # attributes (e.g. the default name/size) that aren't being
        # overridden. openpyxl Fonts are immutable; build a new one
        # with the existing values + overrides.
        existing = cell.font
        cell.font = Font(
            name=font_kwargs.get("name", existing.name),
            size=font_kwargs.get("size", existing.size),
            bold=font_kwargs.get("bold", existing.bold),
            italic=font_kwargs.get("italic", existing.italic),
            color=font_kwargs.get("color", existing.color),
        )

    # Fill (background color).
    bg = _normalize_color(props.get("bg"))
    if bg:
        cell.fill = PatternFill(
            start_color=bg, end_color=bg, fill_type="solid"
        )

    # Borders. Build a Border with every Side independently — Excel's
    # default Border has all sides None, so unset sides stay invisible.
    sides = {
        "top": _border_side(props.get("border_top")),
        "bottom": _border_side(props.get("border_bottom")),
        "left": _border_side(props.get("border_left")),
        "right": _border_side(props.get("border_right")),
    }
    if any(s is not None for s in sides.values()):
        cell.border = Border(
            top=sides["top"],
            bottom=sides["bottom"],
            left=sides["left"],
            right=sides["right"],
        )

    # Alignment.
    align_kwargs: dict = {}
    text_align = props.get("text_align")
    if text_align in ("left", "center", "right"):
        align_kwargs["horizontal"] = text_align
    indent = props.get("indent")
    if isinstance(indent, (int, float)) and indent > 0:
        align_kwargs["indent"] = int(indent)
    if align_kwargs:
        existing = cell.alignment
        cell.alignment = Alignment(
            horizontal=align_kwargs.get(
                "horizontal", existing.horizontal
            ),
            vertical=existing.vertical,
            indent=align_kwargs.get("indent", existing.indent),
        )


def cell_formula(
    var: Variable,
    period_idx: int,
    var_id: str,
    translator: ExcelTranslator,
    current_sheet: str,
    self_address: VariableAddresses,
) -> Optional[str]:
    """The Excel formula this cell will hold — or ``None`` when the
    cell emits a VALUE instead.

    The single answer to "does Python-Excel parity hold here", and the
    only one: a cell falls back to its computed number by four
    different routes (no expression at all; a positional list whose
    item is an authored literal; a translation that raised; a
    render that collapsed a foreign list positionally and would
    therefore compute the wrong number). Anything that needs to KNOW
    whether a cell stays live — the writer itself, or an external
    consumer that shows the workbook's formulas — must ask here rather
    than re-derive it, or the two answers drift and the consumer starts
    claiming formulas the file does not contain.
    """
    expr = var._expr
    # The coordinate this cell belongs to (an expanded subrow, or a
    # head row showing an authored-expression track): tracked operands
    # then resolve to their same-track rows, never to their blend row.
    track_role = getattr(var, '_track_role', None)
    if expr is None:
        # A tracked line without a shared expression may still show an
        # AUTHORED-expression track (``forecast=recurrence(...)``): the
        # display-default row is that track's formula. Same answer for
        # the file (the expansion sets it on the copy) and for an
        # external caller asking about the original line.
        shown_track = getattr(var, 'shown_track', None)
        if callable(shown_track):
            found = shown_track()
            if found is not None:
                shown_role, expr = found
                if track_role is None:
                    track_role = shown_role
    if isinstance(expr, ListExpr):
        # Positional list: item i belongs to period i. A Literal item is an
        # authored value — emit it as a raw input cell, not a ``=0.0``
        # formula; only ref/expression items translate below.
        items = expr.items
        item = items[period_idx] if 0 <= period_idx < len(items) else (
            items[-1] if items else None
        )
        if item is None or isinstance(item, Literal):
            expr = None
    if expr is not None:
        # Variables with an expression tree → translate to Excel formula.
        # Variables without one (plain inputs, lambda recurrence results) fall
        # through to value-emission below.
        try:
            excel = translator.translate(
                expr,
                period_idx,
                current_sheet=current_sheet,
                self_address=self_address,
                self_var=var,
                track_role=track_role,
            )
        except Exception as exc:
            logger.warning(
                "Failed to translate formula for %r (period %d): %s. Falling back to computed value.",
                var_id, period_idx, exc,
            )
            excel = None
        if excel and getattr(translator, 'last_render_lossy', False):
            # The formula collapsed a foreign list positionally — it
            # would compute the WRONG number. The truth wins: emit the
            # computed value.
            excel = None
        if excel and getattr(translator, 'last_render_inlined', False):
            # A coordinate operand's value is baked into the text —
            # true now, stale after that operand is edited. Flag it so
            # a consumer that caches formula text can skip this line.
            translator.inlined_value_seen = True
        if excel:
            excel_str = str(excel)
            if _UNRESOLVED_PLACEHOLDER.search(excel_str):
                logger.warning(
                    "Formula for %r contains unresolved placeholder(s): %s. "
                    "The resulting cell will be invalid in Excel.",
                    var_id, excel_str,
                )
            return excel_str if excel_str.startswith("=") else f"={excel_str}"
    return None


def opening_content(var: Variable, translator: ExcelTranslator, sheet_name: str) -> Any:
    """What ``var`` shows in the opening column — a formula (``=C4-1``, the
    day before the start), a number (``0``, the period before the first)
    or ``None`` for an empty cell. Public so every renderer fills the same
    opening cells with the same content."""
    renderer = getattr(translator, "renderer", None)
    render = getattr(renderer, "render_opening", None)
    if not callable(render):
        return None
    got = render(var, sheet_name)
    if isinstance(got, str):
        return got if got.startswith("=") else f"={got}"
    return got


def opening_value(var: Variable) -> Any:
    """The value an opening cell holds (see :func:`opening_content`): the
    day before the first period for an end row, 0 for a period number —
    of the row itself, or of the row a header mirror repeats."""
    from ...core.time import _first_day, _parse, resolve_default_window, time_field_of
    mirrored = getattr(var, "_mirror_of", None)
    src = var if mirrored is None else mirrored
    field = time_field_of(src)
    if field == "index":
        return 0
    if field != "end":
        return None
    window = resolve_default_window(src)
    if window is None or window.start is None:
        return None
    source = getattr(window, "start_source", None)
    first = getattr(source, "_value", None) if source is not None else None
    if not isinstance(first, date):
        first = (getattr(window, "first_day", None)
                 or _first_day(_parse(window.start, window.grain), window.grain))
    return first - timedelta(days=1)


def title_cells(
    sheet_mv: Any, view: Any, columns: dict, sheet_name: str,
) -> dict[tuple[int, int], tuple[str, str]]:
    """The title row of one sheet (``sheet={'title': True}``): ``{(row,
    col): (text, band role)}`` — the sheet's name over the label column
    (``'upper'`` upper-cases it), its description in the description
    column when one is projected, else just right of the name. Public so
    every renderer writes the same title."""
    setting = getattr(view, "sheet_title", None)
    if not setting:
        return {}
    source = getattr(sheet_mv, "_sheet_source", None)
    if source is None:
        source = sheet_mv
    name = str(getattr(source, "display_name", None) or sheet_name)
    if setting == "upper":
        name = name.upper()
    label = columns.get("label") or 1
    out = {(1, label): (name, "title")}
    note = getattr(source, "description", None)
    if note:
        out[(1, columns.get("description") or label + 1)] = (str(note), "title_note")
    return out


def _cell_content(
    var: Variable,
    period_idx: int,
    var_id: str,
    translator: ExcelTranslator,
    current_sheet: str,
    self_address: VariableAddresses,
) -> Any:
    """The cell's content: its formula when parity holds, else the
    computed value."""
    formula = cell_formula(
        var, period_idx, var_id, translator, current_sheet, self_address,
    )
    if formula is not None:
        return formula

    value = var._value
    if isinstance(value, list):
        if period_idx < len(value):
            return _scalarize(value[period_idx])
        return None
    return _scalarize(value)


def _scalarize(value: Any) -> Any:
    """The cell value for one computed value.

    Numbers, text, booleans and ``None`` go in as they are. A date or
    datetime goes in as a real spreadsheet date (openpyxl stores it as
    the day number), never as its text: text only looks like a date,
    and a comparison against a real date then answers wrong because
    text sorts above every number. A ``timedelta`` goes in as its count
    of days (12 hours is 0.5), the number a formula adds to a date.
    Anything else (custom objects, numpy scalars) is stringified so the
    cell write doesn't raise — readable if imperfect.
    """
    if value is None or isinstance(value, (int, float, str, bool)):
        return value
    if isinstance(value, datetime):
        # A spreadsheet date has no time zone: keep the clock reading.
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return value
    if isinstance(value, timedelta):
        return value / timedelta(days=1)
    return str(value)


def _safe_sheet_name(name: str) -> str:
    """Excel tab names: max 31 chars, can't contain : \\ / ? * [ ], and
    can't begin or end with an apostrophe."""
    forbidden = set(':\\/?*[]')
    cleaned = "".join("_" if c in forbidden else c for c in name)
    return cleaned[:31].strip("'") or "Sheet"


def column_widths(view: Any, columns: dict, last_col: int) -> dict[int, float]:
    """``{column: width}`` the view declares for one sheet, by role
    (``sheet={'columns': [('label', {'width': 40}), …]}``): each role at
    its column (``columns``, the layout's own map), ``period`` over every
    column from the first period to ``last_col``. Public so every
    renderer draws the same widths."""
    out: dict[int, float] = {}
    for role, spec in (getattr(view, "sheet_columns", None) or ()):
        width = spec.get("width")
        if not isinstance(width, (int, float)) or isinstance(width, bool):
            continue
        if role == "period":
            first = columns.get("period")
            cols = range(first, last_col + 1) if first else range(0)
        else:
            col = columns.get(role)
            cols = range(col, col + 1) if col else range(0)
        for c in cols:
            out[c] = float(width)
    return out


def _auto_width_first_column(ws: Worksheet, label_col: int = 1,
                             meta_article: bool = False,
                             meta_unit: bool = False) -> None:
    """Widen the LABEL column so names read; the lead metadata
    columns (No. / unit) stay narrow."""
    max_len = 10
    for cell in ws[get_column_letter(label_col)]:
        if cell.value is not None:
            max_len = max(max_len, len(str(cell.value)))
    ws.column_dimensions[get_column_letter(label_col)].width = min(
        max_len + 2, 50
    )
    if meta_article:
        ws.column_dimensions[get_column_letter(1)].width = 7
    if meta_unit:
        ws.column_dimensions[get_column_letter(label_col + 1)].width = 6
