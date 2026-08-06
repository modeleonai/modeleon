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
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.worksheet import Worksheet

from .translator import ExcelTranslator
from .layout import LayoutEngine, sheet_name_for
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
    translator = ExcelTranslator(addresses, var_to_sheet)

    wb = Workbook()
    # ``Workbook()`` always creates a default sheet — the assert pins
    # that contract for type-checkers (Workbook.active is Optional).
    assert wb.active is not None
    wb.remove(wb.active)

    for sheet_mv, sheet_name in zip(roots, sheet_names):
        ws = wb.create_sheet(title=sheet_name)
        _write_sheet(ws, sheet_mv, addresses, translator, sheet_name,
                     engine.section_header_rows, engine.time_headers)

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


def expand_tracked_tree(root: "MultiVariableBase",
                        mode: str = 'rows') -> "MultiVariableBase":
    """Expand tracked Variables into per-track rows for emission (P3).

    ``mode`` is the view's ``tracks`` setting (§ExcelView): ``'rows'``
    (default) — the display-default row plus labeled per-track rows;
    ``'blend'`` — the display-default row ONLY (the clean board book);
    ``'compare'`` — period-major column groups, refused until that
    layout lands.

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

    if mode == 'compare':
        raise ValueError(
            "ExcelView tracks='compare' (period-major column groups) "
            "is not built yet — use 'blend' or 'rows'."
        )
    if not _has_tracked(root):
        return root

    def _flat_copy(var: Variable, series: object) -> Variable:
        flat = Variable(display_name=var._display_name)
        flat._value = (list(series) if isinstance(series, list) else series)
        flat.var_type = 'list' if isinstance(series, list) else 'scalar'
        flat._unit = var._unit
        flat._start = var._start
        flat._grain = var._grain
        flat._excel_props = dict(var._excel_props)
        src_code = getattr(var, '_source_code', None)
        if src_code:
            flat._source_code = src_code
        return flat

    def _walk(mv: "MultiVariableBase") -> "MultiVariableBase":
        out = MultiVariable(
            mv._display_name, excel_props=dict(mv._excel_props) or None
        )
        for name, child in mv._components.items():
            if isinstance(child, MultiVariableBase):
                setattr(out, name, _walk(child))
                continue
            value = getattr(child, '_value', None)
            if not isinstance(value, TrackValues):
                setattr(out, name, child)
                continue
            decl = resolve_tracks_decl(child)
            blend_name = getattr(getattr(decl, 'blend', None), 'name', None)
            roles = list(value.roles)
            default_role = (blend_name if blend_name in roles else roles[0])
            default = _flat_copy(child, value[default_role])
            if child._expr is not None:
                # The shared formula renders against display-default
                # rows of its operands — the live universe in Excel
                # formulas, exactly the BLEND sheet's law.
                default._set_expr(child._expr)
            setattr(out, name, default)
            if mode == 'blend':
                continue    # the clean board book: default rows only
            for role in roles:
                if role == default_role:
                    continue
                extra = _flat_copy(child, value[role])
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

    **Flat-container exception.** When ``root`` itself carries the
    duck-typed ``_is_context`` marker (set by render-time wrapper
    classes that hold mixed-bag children expected to render as ONE
    sheet — e.g. a focused view's cross-scope binding container),
    the first-depth promotion rule is suppressed: the whole root
    collapses into a single virtual sheet, and every nested MV
    renders as a section inside it. Lets render-layer wrappers
    consolidate what would otherwise be N sibling tabs into one,
    without needing engine-side knowledge of which wrapper produced
    them.
    """
    if getattr(root, "_is_context", False):
        label = root.display_name or root.python_name or "Sheet1"
        items = [(n, root._components[n]) for n in root._component_order]
        return [_make_virtual_sheet(label, items, source=root)]

    direct_vars: list[tuple[str, Variable]] = []
    sub_mvs: list[MultiVariableBase] = []
    for name in root._component_order:
        comp = root._components[name]
        if isinstance(comp, Variable):
            direct_vars.append((name, comp))
        elif isinstance(comp, MultiVariableBase):
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
    virt._excel_props = {'tab': True}
    virt._excel_layout = getattr(source, "_excel_layout", None) if source else None
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


def _write_sheet(
    ws: Worksheet,
    sheet_mv: MultiVariableBase,
    addresses: dict[str, VariableAddresses],
    translator: ExcelTranslator,
    sheet_name: str,
    section_header_rows: dict[str, tuple],
    time_headers: dict[int, tuple],
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

    # Timeline header — period labels across row 1 (cols B onward), one per
    # value column on THIS sheet. Present only when the layout resolved a time
    # window for the sheet (recorded in ``time_headers``); a window-less sheet
    # keeps row 1 for data and renders byte-identically to before.
    header_info = time_headers.get(id(sheet_mv))
    if header_info is not None:
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
                    for cell_addr in addresses[vid].values:
                        _, col_idx = coordinate_to_tuple(cell_addr)
                        header_max_col = max(header_max_col, col_idx)
        n_cols = header_max_col - 1
        if n_cols > 0:
            for i, lbl in enumerate(_period_labels(start, None, n_cols, grain, style)):
                _apply_styling(ws.cell(row=1, column=2 + i, value=lbl), {'bold': True})

    for _, comp in sheet_mv.walk():
        if isinstance(comp, MultiVariableBase):
            section_id = comp.id or comp.python_name or comp._name_in_parent
            pos = section_header_rows.get(section_id)
            if pos is not None and comp.display_name:
                row, col = pos
                header_cell = ws.cell(
                    row=row, column=col, value=comp.display_name
                )
                # Band styling: the resolved view's ``band_styles[role]`` provides
                # defaults UNDER the MV's own excel_props (explicit props win, so
                # existing models are byte-identical). Role derives from the MV's
                # own props — a fill (``bg``) marks a ``section``, otherwise a
                # ``subsection`` — matching the section/subsection convention.
                own = comp._excel_props or {}
                role = "section" if own.get("bg") else "subsection"
                view_bands = resolve_excel_view(comp).band_styles or {}
                band_props = {**view_bands.get(role, {}), **own}
                if band_props:
                    _apply_styling(header_cell, band_props)
                    # Extend the band across the row (label cell already styled).
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

        label = comp.display_name
        # Suffix the unit when the Variable has one — "Revenue ($)".
        unit = getattr(comp, '_unit', None)
        if unit is not None:
            unit_str = str(unit)
            if unit_str:
                label = f"{label} ({unit_str})"
        if label:
            ws[addr.name] = label

        # Optional per-variable number format (e.g. dates → "mmm yyyy").
        # Two override sources, checked in priority order:
        #   1. ``_excel_layout.format`` — structured layout tag from
        #      pro's ``ExcelLayout`` dataclass. Wins because it's the
        #      explicit per-Variable format on the layout system.
        #   2. ``_excel_props['number_format']`` — legacy flag-dict path.
        # Engine reads the attribute opaquely (no pro import); the
        # layout type lives in pro.
        number_format = None
        layout = getattr(comp, "_excel_layout", None)
        if layout is not None:
            layout_format = getattr(layout, "format", None)
            if layout_format:
                number_format = str(layout_format)
        if number_format is None and comp._excel_props:
            number_format = comp._excel_props.get("number_format")
        # Date-valued Variables (literals or formulas like ``=EDATE(...)``)
        # compute an Excel date serial; without a date number-format the cell
        # shows the raw serial (44562) instead of the date. Default to ISO
        # date display unless an explicit format was provided.
        if number_format is None and getattr(comp, "value_type", None) == "datetime":
            number_format = "yyyy-mm-dd"

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
        # Apply the same styling to the label cell so the row reads as
        # one styled unit — when a user marks a Variable bold, both the
        # label and its values bold (matches Excel users' expectations
        # that styling describes the "row" not individual cells).
        if label and props:
            _apply_styling(ws[addr.name], props)

    _auto_width_first_column(ws)
    # Freeze the label column so row labels stay visible while scrolling across
    # many period columns — standard for wide financial models. When a timeline
    # header occupies row 1, freeze it too (``B2``) so the period labels stay
    # pinned alongside the row labels.
    ws.freeze_panes = "B2" if id(sheet_mv) in time_headers else "B1"


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


def _cell_content(
    var: Variable,
    period_idx: int,
    var_id: str,
    translator: ExcelTranslator,
    current_sheet: str,
    self_address: VariableAddresses,
) -> Any:
    """Return the value to put in a cell: an Excel formula for computed variables, else the raw value."""
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
        if excel:
            excel_str = str(excel)
            if _UNRESOLVED_PLACEHOLDER.search(excel_str):
                logger.warning(
                    "Formula for %r contains unresolved placeholder(s): %s. "
                    "The resulting cell will be invalid in Excel.",
                    var_id, excel_str,
                )
            return excel_str if excel_str.startswith("=") else f"={excel_str}"

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


def _auto_width_first_column(ws: Worksheet) -> None:
    """Widen column A so labels are readable."""
    max_len = 10
    for cell in ws[get_column_letter(1)]:
        if cell.value is not None:
            max_len = max(max_len, len(str(cell.value)))
    ws.column_dimensions[get_column_letter(1)].width = min(max_len + 2, 50)
