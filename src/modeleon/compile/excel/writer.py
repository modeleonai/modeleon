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
from pathlib import Path
from typing import Any, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.worksheet import Worksheet

from .translator import ExcelTranslator
from .layout import LayoutEngine, effective_hidden, sheet_name_for
from .addresses import VariableAddresses
from .view import resolve_excel_view
from ...core.expr import ListExpr, Literal
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable

logger = logging.getLogger(__name__)

_UNRESOLVED_PLACEHOLDER = re.compile(r"\{[A-Z_][A-Z0-9_]*\}")


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
    from .view import resolve_excel_view
    track_mode = resolve_excel_view(root).tracks or 'rows'
    root = expand_tracked_tree(root, mode=track_mode)
    # (the compare SELECTION arrives only through the pro export door;
    # a declared tracks='compare' here shows every declared track)
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
    engine = LayoutEngine(roots, flatten_nested_sheets=True)
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

    var_to_sheet = _build_var_to_sheet(addresses, roots, sheet_names)
    identity_cells = _identity_homes(roots, addresses)
    translator = ExcelTranslator(addresses, var_to_sheet,
                                 identity_cells=identity_cells)

    wb = Workbook()
    # ``Workbook()`` always creates a default sheet — the assert pins
    # that contract for type-checkers (Workbook.active is Optional).
    assert wb.active is not None
    wb.remove(wb.active)

    for sheet_mv, sheet_name in zip(roots, sheet_names):
        ws = wb.create_sheet(title=sheet_name)
        # Служебный таб (страница-виджет, аналитика карточек): лист
        # пишется целиком — значения и формулы живут, аудит возможен —
        # но вкладка нативно скрыта.
        if effective_hidden(sheet_mv):
            ws.sheet_state = "hidden"
        _write_sheet(ws, sheet_mv, addresses, translator, sheet_name,
                     engine.section_header_rows, engine.time_headers,
                     engine_meta=engine.meta_by_sheet,
                     board_var_ids=engine.board_var_ids,
                     section_header_echoes=engine.section_header_echoes,
                     inline_section_ids=engine.inline_section_ids,
                     period_start=engine.period_start_by_sheet.get(
                         id(sheet_mv), 0),
                     period_var_ids=engine.period_var_ids,
                     meta_fields=engine.meta_fields_by_sheet.get(
                         id(sheet_mv), ()),
                     constants_caption=engine.constants_by_sheet.get(
                         id(sheet_mv)),
                     totals_plan=engine.totals_by_sheet.get(sheet_name),
                     group_plan=engine.groups_by_sheet.get(sheet_name),
                     row_outline=engine.row_outline_by_sheet.get(sheet_name))
        _plan = engine.totals_by_sheet.get(sheet_name)
        if _plan is not None:
            from .totals_writer import write_totals

            write_totals(
                ws, sheet_mv, _plan, addresses, var_to_sheet, sheet_name,
            )
        # The blend row references its track subrows instead of baking
        # the splice — fill a fact in the file and «live» moves.
        from .blend_writer import write_blend_links

        write_blend_links(
            ws, sheet_mv, addresses, var_to_sheet, sheet_name,
        )
        # …and a DERIVED track subrow explains itself in its own
        # coordinates, instead of sitting as a literal under a row that
        # does. Runs after the blend pass: the blend row is its own
        # case and must not be overwritten here.
        from .track_writer import write_track_formulas

        write_track_formulas(
            ws, sheet_mv, addresses, var_to_sheet, sheet_name,
        )

    wb.save(path)


def _has_tracked(mv: "MultiVariableBase") -> bool:
    from ...core.tracks import TrackValues
    for child in mv._components.values():
        if isinstance(child, MultiVariableBase):
            if _has_tracked(child):
                return True
        elif isinstance(getattr(child, '_value', None), TrackValues):
            return True
    return False


#: Metric columns of the compare lens, in REGISTRY order — the same
#: catalogue the IDE's grid draws, so the file and the view cannot
#: disagree on what «откл %» means. Values are FRACTIONS wearing a
#: percent number format (Excel's own semantics); the grid multiplies
#: by 100 for display of the same numbers.
_METRIC_REGISTRY = ("откл", "откл%", "уд.вес", "%роста")
_METRIC_LABELS = {
    "откл": "откл", "откл%": "откл %",
    "уд.вес": "уд. вес", "%роста": "% роста",
}
_METRIC_FORMATS = {"откл%": "0.0%", "уд.вес": "0.0%", "%роста": "0%"}
#: Which metrics need the given/follow PAIR visible — without it the
#: number has no meaning and the column drops (the grid's own law).
_METRIC_NEEDS_PAIR = {"откл", "откл%"}


def _resolve_metric_keys(selection, pair_available) -> list[str]:
    """The grid's ``resolveMetrics``, mirrored: ``None`` = the
    historical «откл» default, ``[]`` = explicitly none; registry
    order regardless of the checkbox order."""
    keys = ["откл"] if selection is None else list(selection)
    return [
        k for k in _METRIC_REGISTRY
        if k in keys and (k not in _METRIC_NEEDS_PAIR or pair_available)
    ]


def _section_total_of(child):
    """The denominator line for «уд. вес»: the nearest ``итого``
    sibling at or above this line's section — the same qpath
    convention the grid's ``sectionTotalOf`` walks. ``None`` when no
    ancestor section carries one (the metric stays honestly empty)."""
    from ...core.multi_variable import MultiVariableBase

    name = getattr(child, '_name_in_parent', None)
    node = getattr(child, '_owner', None) or getattr(child, '_parent', None)
    # A line NAMED «итого» measures against the section above its own.
    if name == 'итого' and node is not None:
        node = (getattr(node, '_parent', None)
                or getattr(node, '_owner', None))
    while node is not None:
        if isinstance(node, MultiVariableBase):
            total = node._components.get('итого')
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
        #: role → slot. Visible tracks first (the lens's own order),
        #: metric columns next (registry order, slots ``metric_base +
        #: i``), hidden blend operands last — so the VISIBLE part of a
        #: group is exactly what the grid draws, and the file-only
        #: extras trail it.
        self.slot_of = slot_of
        self.hidden = hidden               # set of hidden ROLE names
        self.given = given
        self.follow = follow
        self.metric_keys = metric_keys
        self.metric_base = metric_base
        self.stride = stride
        #: (minuend, subtrahend) for the pair metrics — ANY two
        #: visible columns, the grid's own generalized rule: the
        #: subtrahend is the план-like column (literal «план», else
        #: the follow), the minuend факт-like (given, else the blend,
        #: else the first other column). The rigid blend pair hid
        #: «откл» on факт+план — Arrual's flagship view.
        self.pair = pair


def _compare_plan(decl, selection, roles,
                  metrics=None) -> Optional[_ComparePlan]:
    """Slots for one period group, mirroring the IDE lens's own order:
    given, follow, then the rest of the selection; then the METRIC
    columns («откл», «уд. вес», …); the blend and both its sources
    always ride (hidden when unchecked) so the splice formula keeps
    its operands. Pair metrics need both blend tracks CHECKED — a
    deviation between anything else is a number without a meaning
    (the grid's own law)."""
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
        sub = (
            'план' if 'план' in visible and given != 'план'
            else follow if follow in visible
            else visible[1]
        )
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
    """One metric COLUMN as a live Variable — the same four the grid
    draws, each an expression the file can hold:

    * «откл»    = given − follow (engine arithmetic, sums honestly);
    * «откл %»  = (given − follow) / ABS(follow), a FRACTION wearing a
      percent format — Excel's own percent semantics;
    * «уд. вес» = the line's shown series over the nearest section
      «итого»'s — AFTER aggregation by construction, because both
      operands' bucket cells are already aggregates;
    * «% роста» = shown over the same cell a year earlier, cell by
      cell (a ListExpr: early periods have no prior year and stay
      honestly EMPTY, exactly the grid's null).

    ``None`` when the metric's ingredients are missing on this line —
    the column simply stays empty there, never a zero."""
    from ...core.expr import BinOp, FuncCall, Literal, ListExpr, Subscript, VarRef

    def _div(a, b):
        return (a / b) if (a is not None and b not in (None, 0)) else None

    shown_vals = value[default_role]
    n = len(shown_vals)

    if key in ("откл", "откл%"):
        if plan.pair is None:
            return None
        g = copies.get(plan.pair[0])
        f = copies.get(plan.pair[1])
        if g is None or f is None:
            return None
        if key == "откл":
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

    if key == "уд.вес":
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

    if key == "%роста":
        from ...core.time import resolve_default_window
        window = resolve_default_window(child)
        grain = getattr(window, 'grain', None) or 'month'
        steps = {'month': 12, 'quarter': 4, 'year': 1}.get(grain)
        if not steps or n <= steps:
            return None
        items = []
        for t in range(n):
            if t < steps:
                items.append(Literal(value=None))
                continue
            items.append(BinOp(
                op='/',
                left=Subscript(base=VarRef(var=child), key=t),
                right=Subscript(base=VarRef(var=child), key=t - steps),
            ))
        vals = [
            None if t < steps else _div(shown_vals[t], shown_vals[t - steps])
            for t in range(n)
        ]
        out = Variable(display_name='')
        out._value = vals
        out.var_type = 'list'
        out._set_expr(ListExpr(items=items))
        # A prior-year BUCKET does not exist yet (the grid's own null)
        # — the totals writer leaves these cells empty instead of
        # projecting a lag whose step no longer matches the grain.
        out._metric_no_bucket = True
        return out

    return None


def expand_tracked_tree(root: "MultiVariableBase",
                        mode: str = 'rows',
                        tracks: Optional[list[str]] = None,
                        metrics: Optional[list[str]] = None,
                        ) -> "MultiVariableBase":
    """Expand tracked Variables into per-track rows for emission (P3).

    ``mode`` is the view's ``tracks`` setting (§ExcelView): ``'rows'``
    (default) — the display-default row plus labeled per-track rows;
    ``'blend'`` — the display-default row ONLY (the clean board book);
    ``'compare'`` — period-major column GROUPS: one column per track
    under each period, the same drawing the IDE's track lens makes.
    ``tracks`` narrows compare to a selection (the lens's checkboxes);
    the blend and both its source tracks are always emitted so the
    splice formula stays honest — unchecked ones ride as HIDDEN
    columns rather than silently degrading the blend to one operand.
    An all-scalar tracked line (a record's card field) stays a single
    cell in compare mode: a scalar has no period columns to group.

    The BLEND default sheet (§16.9): a tracked line keeps its NAME and
    qpath on a row carrying its display-default series — the
    synthesized live track when the model blends, else the first
    declared track — so every formula that references the line renders
    against that row (Excel parity of the live universe). The other
    tracks follow as sibling value rows labeled «name · track-label».

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
    # The drawing the IDE's track lens already makes, now in the file.
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
        # ``_components`` / ``_component_order`` surfaces are derived
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
            if (getattr(child, '_role_kwargs', None) is not None
                    or child._expr is None):
                default._blend_is_splice = True
            if default_role in (getattr(child, '_role_kwargs', None) or {}):
                # The shown series was entered, not derived — mark it
                # so consumers can distinguish the two.
                default._authored_track = default_role
            if child._expr is not None:
                # The shared formula renders against display-default
                # rows of its operands — the live universe in Excel
                # formulas, exactly the BLEND sheet's law.
                default._set_expr(child._expr)
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
                for mi, mkey in enumerate(plan.metric_keys):
                    mvar = _metric_variable(
                        mkey, child, default_role, value,
                        copies, plan, decl,
                    )
                    if mvar is None:
                        continue
                    mvar._display_name = (
                        f"{child._display_name or name} "
                        f"· {_METRIC_LABELS[mkey]}"
                    )
                    mvar._col_slot = plan.metric_base + mi
                    mvar._col_head = default
                    mvar._col_word = _METRIC_LABELS[mkey]
                    fmt = _METRIC_FORMATS.get(mkey)
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
    like records is the case that needs it: focusing a roster of
    twenty-nine people used to open twenty-nine tabs, one per person,
    when the thing the author declared was one list.

    A child that asks for a tab OF ITS OWN still gets one — the marker
    sets the default for the container's children, it does not overrule
    a child that declared otherwise. So one landmark record can stand
    apart while the rest stay a list.

    ``_is_context`` keeps working as a synonym: render-time wrappers
    (a focused view's cross-scope binding container) were the first
    users of this rule, back when it was theirs alone.
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
    # (a collection) reached the layout as a plain MultiVariable and
    # was stacked section-per-record instead of laid out as the table
    # it declares itself to be. Styling props stay on the source's
    # cells; only orientation describes the SHEET.
    virt._excel_props = {'tab': True}
    _src_props = getattr(source, "_excel_props", None) or {}
    if _src_props.get('orient'):
        virt._excel_props['orient'] = _src_props['orient']
    # Служебность едет с источником: страница-виджет, завёрнутая в
    # виртуальный таб, остаётся СКРЫТЫМ листом книги (строки-подписи
    # виджетов — не финансовые строки).
    if effective_hidden(source):
        virt._excel_props['hidden'] = True
    virt._excel_layout = getattr(source, "_excel_layout", None) if source else None
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
    # could not answer. A collection wraps itself this way, so every
    # tracked line on a records sheet lost its splice — the workbook
    # showed the base track where the app showed the blend.
    from ...core.tracks_decl import resolve_tracks_decl
    _src_decl = resolve_tracks_decl(source) if source is not None else None
    if _src_decl is not None:
        virt.tracks = _src_decl
        # …and the WINDOW with it: the splice asks the sheet where the
        # close boundary falls, and an unanswerable window read as
        # "no boundary at all" — every cell took the before-boundary
        # form, so the periods AFTER the close spliced backwards.
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
        while name in seen:
            tag = str(suffix)
            name = f"{base[:31 - len(tag)]}{tag}"
            suffix += 1
        seen.add(name)
        names.append(name)
    return names


def _identity_homes(roots: "list[MultiVariableBase]",
                    addresses: "dict[str, VariableAddresses]"):
    """The emission tree's identity map (see ``identity_cells_for``).

    Homes are the rows whose cells genuinely HOLD the series: track
    subrows and untracked rows. The blend head of a tracked line is
    excluded — its cells are the splice, and a план formula referencing
    a blend cell would move the moment a fact lands in the file.
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


def resolve_number_format(comp: Variable) -> str | None:
    """The cell number format a Variable's values display with — one
    resolution shared by the .xlsx writer and any grid mirroring it.

    Priority: the Variable's own ``_excel_layout.format`` → its own
    ``excel_props['number_format']`` → the nearest ancestor MV's
    ``excel_props['number_format']`` (a section states its unit
    convention once) → the datetime default → the resolved ExcelView's
    ``formats['number']`` floor for numeric cells.
    """
    layout = getattr(comp, "_excel_layout", None)
    if layout is not None:
        layout_format = getattr(layout, "format", None)
        if layout_format:
            return str(layout_format)
    props = getattr(comp, "_excel_props", None)
    if props and props.get("number_format"):
        return props["number_format"]
    node = getattr(comp, "_owner", None)
    while node is not None:
        props_up = getattr(node, "_excel_props", None)
        if props_up and props_up.get("number_format"):
            return props_up["number_format"]
        node = getattr(node, "_parent", None)
    if getattr(comp, "value_type", None) == "datetime":
        return "yyyy-mm-dd"
    if getattr(comp, "value_type", None) in ("int", "float"):
        fmts = resolve_excel_view(comp).number_formats
        if fmts:
            # Unit-keyed first: ``formats={'₸': …, '%': …}`` — declare
            # the unit on the line, the format follows it (and follows
            # DERIVED lines through unit algebra). ``'number'`` is the
            # generic floor for unitless numerics.
            unit = getattr(comp, "_unit", None)
            if unit is not None:
                by_unit = fmts.get(str(unit))
                if by_unit:
                    return by_unit
            return fmts.get("number")
    return None


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
    config = _inherited_format_by_type(comp)
    if not config:
        return None
    # ``format_by_type=True`` uses the resolved view's palette (seeded by the
    # house default ``mo.default_excel_view.type_colors``); a dict value is a
    # legacy inline palette and is used as-is.
    colors = config if isinstance(config, dict) else (
        resolve_excel_view(comp).type_colors or {}
    )
    return colors.get(_cell_type(comp, sheet_name, var_to_sheet))


def lead_cells_for_row(
    comp: object,
    row: int,
    label_col: int,
    *,
    article: bool,
    unit: bool,
    fields: tuple = (),
    sum_resolver: object = None,
) -> dict[tuple[int, int], str]:
    """The lead-column cells of ONE row: ``{(row, col): text}``.

    The article number in column A, the unit in its own column right
    after the label, then one column per projected field in declared
    order. Shared with the run-state builder so the grid and the
    workbook place the same text in the same cells — the geometry
    lives here once instead of being re-derived wherever it is
    needed, which is how the two surfaces drift apart.

    A ``('row_sum', caption)`` field is COMPUTED, not read: rows that
    opt in via ``excel_props={'row_sum': True}`` get whatever
    ``sum_resolver(comp)`` synthesizes — the writer passes a live
    ``=SUM(...)`` over the row's value cells, the run-state builder
    the computed total — placed at the field's own column so both
    surfaces agree on geometry by construction.
    """
    out: dict[tuple[int, int], str] = {}
    if article:
        value = (getattr(comp, "_excel_props", None) or {}).get("article")
        if value:
            out[(row, 1)] = str(value)
    if unit:
        value = getattr(comp, "_unit", None)
        if value is not None:
            out[(row, label_col + 1)] = str(value)
    first = label_col + 1 + (1 if unit else 0)
    for i, (attr, _caption) in enumerate(fields or ()):
        if attr == "row_sum":
            flagged = (getattr(comp, "_excel_props", None) or {}).get(
                "row_sum"
            )
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


def row_sum_formula(refs: "list[str]") -> "str | None":
    """``=SUM(...)`` over a row's value cells, ranges compacted.

    The cells may be non-contiguous (subtotal columns interleave with
    the months), so consecutive same-row runs compress to ``B5:D5``
    pieces joined by commas — the formula a hand-built book would
    carry on the row's «Итого» column.
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

    One law with the grid's ``_row_sum_display`` refusals: a SCALAR
    (single value cell) is not a run of periods — printing «=SUM(D5)»
    beside a number the row already shows is noise; an ALL-BLANK row
    would total a fabricated 0 where the grid honestly stays silent.
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
) -> None:
    """Fill a worksheet: an optional timeline header row, section-header
    labels for nested MVs, then a row per Variable (label + per-period
    values or formulas)."""
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
        n_cols = header_max_col - _hdr_first + 1 if _hdr_first else 0
        if n_cols > 0:
            # Header band: the resolved view's ``bands['header']`` styles
            # the period-label row (brand fill, font colour) over the
            # bold floor — the book's masthead, declared once.
            _hdr_bands = (resolve_excel_view(sheet_mv).band_styles or {})
            _hdr_props = {'bold': True, **_hdr_bands.get('header', {})}
            _hdr_start = _hdr_first
            if totals_plan is not None:
                # HIERARCHICAL header — «год» over «кварталы» over
                # «месяцы»: each bucket is a merged span over its
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
                # ± over the columns, and «свернуть 2027» works in
                # the downloaded file exactly as in the grid. Summary
                # RIGHT of detail — where our totals stand.
                ws.sheet_properties.outlinePr.summaryRight = True
                _kinds_depth = {"quarter": 0, "year": 0}
                _enclosers = [k for k in ("year", "quarter")
                              if k in _plan.buckets]
                # A period is a GROUP of ``stride`` columns and the
                # bracket must cover all of them — a level on the
                # first column alone drew Excel brackets over one
                # column in four («группировки непонятные», live).
                # Hidden operand columns go one level deeper: their
                # own collapsed subgroup, so expanding a quarter does
                # not unhide what the lens deliberately hid.
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
                        # centered «2026» floats far from the row
                        # names it belongs to (founder call). The
                        # month row stays centered under it.
                        _align = (
                            'center' if cs["colspan"] == 1 else 'left'
                        )
                        _apply_styling(
                            ws.cell(
                                row=hrow["row"],
                                column=cs["col"],
                                value=cs["label"],
                            ),
                            {**_hdr_props, 'bold': True,
                             'text_align': _align},
                        )
                        _rspan = cs.get("rowspan", 1)
                        if cs["colspan"] > 1 or _rspan > 1:
                            ws.merge_cells(
                                start_row=hrow["row"],
                                start_column=cs["col"],
                                end_row=hrow["row"] + _rspan - 1,
                                end_column=cs["col"] + cs["colspan"] - 1,
                            )
            else:
                _stride = group_plan[0] if group_plan else 1
                _n_periods = max(1, n_cols // _stride)
                for i, lbl in enumerate(
                    _period_labels(start, None, _n_periods, grain, style)
                ):
                    _col0 = _hdr_start + i * _stride
                    _apply_styling(
                        ws.cell(row=1, column=_col0, value=lbl),
                        _hdr_props,
                    )
                    if _stride > 1:
                        ws.merge_cells(
                            start_row=1, start_column=_col0,
                            end_row=1, end_column=_col0 + _stride - 1,
                        )
            # Track column groups: the word row under the period
            # labels — «план | факт/прогноз | …» over every group, the
            # same words the grid's lens paints. Unchecked-but-needed
            # columns (a blend operand the splice formula references)
            # hide themselves rather than widen the visible book.
            if group_plan is not None:
                _g_stride, _g_words, _g_hidden = group_plan
                _word_row = (
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
            _cap_row = (
                totals_plan.depth if totals_plan is not None else 1
            )
            # Lead metadata columns get their own headers — a column
            # nobody labelled is a column the reader has to guess at,
            # and the period labels deliberately start AFTER them.
            # The flag IS the caption: ``meta={'unit': 'Ед. изм.'}``
            # heads the column with that text, while a bare ``True``
            # keeps the historical headerless column. Written only
            # here, inside the timeline block: on a sheet with no
            # period header the layout puts DATA in row 1, and a
            # caption there would land on top of it.
            if isinstance(_meta_article, str) and _meta_article:
                _apply_styling(
                    ws.cell(row=_cap_row, column=1, value=_meta_article),
                    _hdr_props
                )
            if isinstance(_meta_label, str) and _meta_label:
                _apply_styling(
                    ws.cell(
                        row=_cap_row,
                        column=1 + (1 if _meta_article else 0),
                        value=_meta_label,
                    ),
                    _hdr_props,
                )
            # The constants column's header — «Значение» in the book
            # form. Sits immediately left of the first period, which
            # is exactly where an unlabeled run of scalars was read as
            # the sheet having slipped against its own timeline.
            if isinstance(constants_caption, str) and constants_caption \
                    and _hdr_first > 1:
                _apply_styling(
                    ws.cell(row=_cap_row, column=_hdr_first - 1,
                            value=constants_caption),
                    _hdr_props,
                )
            _fcap = 2 + (1 if _meta_article else 0) + (1 if _meta_unit else 0)
            for _i, (_attr, _cap) in enumerate(meta_fields):
                if isinstance(_cap, str) and _cap:
                    _apply_styling(
                        ws.cell(row=_cap_row, column=_fcap + _i, value=_cap),
                        _hdr_props,
                    )
            if isinstance(_meta_unit, str) and _meta_unit:
                _apply_styling(
                    ws.cell(
                        row=_cap_row,
                        column=2 + (1 if _meta_article else 0),
                        value=_meta_unit,
                    ),
                    _hdr_props,
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
                # ``subsection``. (The old rule — own ``bg`` marks a section —
                # made ``bands.section.bg`` unreachable: whoever had a fill kept
                # their own, whoever hadn't was a subsection, so a view could
                # never paint an unstyled разд.)
                own = comp._excel_props or {}
                if _meta_article and own.get('article'):
                    ws.cell(row=row, column=1, value=str(own['article']))
                role = "section" if _hdr_level == 0 else "subsection"
                view_bands = resolve_excel_view(comp).band_styles or {}
                band_props = {**view_bands.get(role, {}), **own}
                if band_props:
                    _apply_styling(header_cell, band_props)
                    # Extend the band across the row (label cell already
                    # styled) — unless the label SHARES its row with
                    # data, as a record's does in a records-oriented
                    # table: there the banner would paint the numbers.
                    if section_id not in (inline_section_ids or set()):
                        for c in range(col + 1, sheet_max_col + 1):
                            _apply_styling(ws.cell(row=row, column=c), band_props)

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
        # unit in its own column right after the label.
        if addr.name:
            _lrc = coordinate_to_tuple(addr.name)
            for (_r, _c), _text in lead_cells_for_row(
                comp,
                _lrc[0],
                _lrc[1],
                article=bool(_meta_article),
                unit=bool(_meta_unit),
                fields=meta_fields,
                # The book's row total is a LIVE formula over the row's
                # own value cells — openpyxl stores an ``=``-string as
                # a formula, so the .xlsx recalculates it on edit.
                # Eligibility mirrors the grid's refusals (scalars,
                # all-blank rows) so the two surfaces stay one drawing.
                sum_resolver=lambda _c_, _a=addr: (
                    row_sum_formula(list(_a.values))
                    if row_sum_eligible(_c_, list(_a.values)) else None
                ),
            ).items():
                ws.cell(row=_r, column=_c, value=_text)
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
        #   1. ``_excel_layout.format`` — structured layout tag from
        #      pro's ``ExcelLayout`` dataclass. Wins because it's the
        #      explicit per-Variable format on the layout system.
        #   2. ``_excel_props['number_format']`` — legacy flag-dict path.
        # Engine reads the attribute opaquely (no pro import); the
        # layout type lives in pro.
        number_format = resolve_number_format(comp)

        props = comp._excel_props or {}
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
        type_color = _type_color_for(comp, sheet_name, translator.var_to_sheet)
        if type_color and "font_color" not in props:
            value_props = {**props, "font_color": type_color}

        for i, cell_addr in enumerate(addr.values):
            cell = ws[cell_addr]
            cell.value = _cell_content(comp, i, var_id, translator, sheet_name, addr)
            if number_format:
                cell.number_format = number_format
            _apply_styling(cell, value_props)
        # Value-driven fill (``bg_nonzero``): a LIVE conditional rule —
        # nonzero NUMBERS wear the color, zeros, blanks and text stay
        # bare (the grid's own law — CellIs «notEqual 0» would paint
        # error tokens and any text) — so the band repaints itself
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

    _auto_width_first_column(
        ws,
        label_col=1 + (1 if _meta_article else 0),
        meta_article=_meta_article,
        meta_unit=_meta_unit,
    )
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
    _freeze_row = 1 + (_hdr_depth if id(sheet_mv) in time_headers else 0)
    ws.freeze_panes = f"{_freeze_col}{_freeze_row}"
    # Native ROW groups mirroring the section tree — the grid's
    # collapsible outline, in the file. The section header is the
    # group's summary and sits ABOVE it, so the ± lands on the header
    # row (summaryBelow=False), exactly where the grid puts its ±.
    if row_outline:
        ws.sheet_properties.outlinePr.summaryBelow = False
        for _r, _lv in row_outline.items():
            ws.row_dimensions[_r].outline_level = min(7, _lv)
    # Служебные строки (``excel_props={'hidden': True}``, наследуется
    # по дереву): написаны полностью — значения, формулы, подписи —
    # но нативно скрыты. Книга остаётся аудируемой; глаз не мусорится.
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

    The single answer to «does Python-Excel parity hold here», and the
    only one: a cell falls back to its computed number by four
    different routes (no expression at all; a positional list whose
    item is an authored literal; a translation that raised; a
    render that collapsed a foreign list positionally and would
    therefore compute the wrong number). Anything that needs to KNOW
    whether a cell stays live — the writer itself, an IDE mirroring
    the workbook — must ask here rather than re-derive it, or the two
    answers drift and the mirror starts claiming formulas the file
    does not contain.
    """
    expr = var._expr
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
    """Coerce unsupported types to ``str`` before handing to openpyxl.

    openpyxl only knows how to serialize ``None`` / ``int`` / ``float`` /
    ``str`` / ``bool`` directly. Anything else (``datetime.date``, custom
    objects, numpy scalars) gets stringified so the cell write doesn't
    raise — the cell shows the ``repr``, which is readable if imperfect.
    """
    if value is None or isinstance(value, (int, float, str, bool)):
        return value
    return str(value)


def _safe_sheet_name(name: str) -> str:
    """Excel tab names: max 31 chars, can't contain : \\ / ? * [ ]."""
    forbidden = set(':\\/?*[]')
    cleaned = "".join("_" if c in forbidden else c for c in name)
    return cleaned[:31] if len(cleaned) > 31 else cleaned


def _auto_width_first_column(ws: Worksheet, label_col: int = 1,
                             meta_article: bool = False,
                             meta_unit: bool = False) -> None:
    """Widen the LABEL column so names read; the lead metadata
    columns (№ / ед.) stay narrow."""
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
