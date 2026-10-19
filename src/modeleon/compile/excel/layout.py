# SPDX-License-Identifier: Apache-2.0
"""Layout engine — walks an MV tree and assigns Excel cell addresses.

Pipeline step 1 of 3 (layout → translator → writer). Takes a list of
top-level MultiVariables, walks the tree depth-first, and produces a
``{var_id: VariableAddresses}`` map the writer uses to emit formulas
and values into the correct cells.

Sheets (MVs with ``_excel_props={'tab': True}``) get their own tab. Nested MVs
become visual sections — a header row plus their children. Variables
get one row each (label column + per-period value columns).

The :class:`VariableAddresses` dataclass lives alongside in
:mod:`modeleon.compile.excel.addresses` — the contract between this
module and its consumers (writer, translator).
"""

from typing import Any, Container, Dict, Iterable, Iterator, List, Optional, Sequence, Union

from .addresses import VariableAddresses
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable


def row_kind(comp: object) -> str:
    """What a line IS, for its look (``bands['input'|'link'|'formula']``)
    and its number: ``'input'`` — typed values; ``'link'`` — a bare
    reference to another line (``b.x = a.x.copy()``, the import that opens
    a calculation block); ``'formula'`` — anything else computed. Whether
    a line is its block's result is the model's word, not its shape: the
    ``excel_props={'result': True}`` mark."""
    if not getattr(comp, 'is_formula', False):
        return 'input'
    from ...core.expr import MethodCall, VarRef
    expr = getattr(comp, '_expr', None)
    if isinstance(expr, MethodCall) and expr.method == 'copy':
        expr = expr.base
    return 'link' if isinstance(expr, VarRef) else 'formula'


def _is_presentation(comp: object) -> bool:
    """A nested ExcelView is a named view recipe — presentation
    metadata that layout must never emit as rows or sections. (The
    ``default_excel_view`` slot bypasses adoption entirely; this guard
    covers the ADOPTED named views.)"""
    from .view import ExcelView

    return isinstance(comp, ExcelView)

LayoutRoot = Union[Variable, MultiVariableBase]


def effective_hidden(node: object) -> bool:
    """True when the node or ANY ancestor carries ``excel_props={'hidden': True}``.

    A hidden row is not a skipped one: it is laid out, computed and
    written with its formulas and labels, so every reference to it
    keeps a real address. The writer hides it with Excel's native
    hidden-row flag. The cascade ``_owner`` → ``_parent`` is the same
    one ``resolve_excel_view`` walks.
    """
    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        if (getattr(n, '_excel_props', None) or {}).get('hidden'):
            return True
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return False


def hidden_key_caption_rows(sheet_mv: Any, key_caption_rows: Dict[str, int]) -> set:
    """The caption rows of keyed lines (``LayoutEngine.key_caption_rows``)
    a renderer hides along with the lines beneath them: every line that
    shares the row is hidden (:func:`effective_hidden`). A visible
    sibling keeps the row for both. Public so the workbook and every
    other renderer hide the same rows."""
    owners: Dict[int, List[bool]] = {}
    for _, comp in sheet_mv.walk():
        if isinstance(comp, Variable) and comp.id in key_caption_rows:
            owners.setdefault(key_caption_rows[comp.id], []).append(effective_hidden(comp))
    return {row for row, hidden in owners.items() if all(hidden)}


def _own_orient(mv: Any) -> Optional[str]:
    """The layout orientation ``mv`` DECLARES for itself.

    Read from the MV's own ``ExcelView`` — the one place a printed form
    is described, so a container of like records that wants to print as
    a table says so the same way it says anything else about its form.
    The ``excel_props={'orient': …}`` spelling still wins where it is
    set.

    Deliberately NOT the cascade. ``resolve_excel_view`` inherits
    fields down the tree, which is right for ink — a font declared at
    the model reaches every sheet — and wrong for geometry: a model
    that declared ``orient='records'`` would turn every section into a
    record table, and a section whose children are Variables has no
    records to lay out, so its rows would silently disappear. A shape
    this violent is stated where it applies.
    """
    own = (getattr(mv, '_excel_props', None) or {}).get('orient')
    if own:
        return str(own)
    view = getattr(mv, 'default_excel_view', None)
    snapshot = getattr(view, 'snapshot', None)
    if snapshot is None:
        return None
    # A view is a real tree — its fields are Variables. ``snapshot()``
    # is the flattened, plain-valued form every other reader uses.
    orient = snapshot().orient
    return str(orient) if orient else None


def number_lines(
    sheet_mv: Any,
    rows: Iterable[str],
    scheme: Any,
    sections: Optional[Container[str]] = None,
) -> Dict[str, str]:
    """``{id: number}`` for one sheet's sections and lines, by ``scheme`` —
    ``'section'``: an outline in the order the sheet stands. A first-level
    section is ``1``, ``2``, …, a section inside it ``1.1``, and a line is
    numbered one level below its section (``1.1.1``); sections and lines
    of one container count together, so a line on the sheet itself is
    ``1``, ``2``, … beside the first-level sections.

    ``rows`` are the ids of the lines that head a row of their own (a
    label cell): a line riding another row — a header-row total, a
    track's slot copy, a matrix's sibling instance — carries no number.
    ``sections`` are the ids of the sections that head a row (the
    layout's ``section_header_rows``); a container laid out elsewhere, a
    nested sheet, is skipped. A link (a line that only repeats another)
    carries no number and takes no place; a line or a section with an
    ``article`` of its own keeps it and takes none, and that section's
    lines count from its article. A table (``orient='records'``) is one
    item. The lines are read from ``sheet_mv`` as they stand now, so
    every caller numbers them by the same rule — the one rule every
    renderer numbers by."""
    out: Dict[str, str] = {}
    if scheme != 'section':
        return out
    _number_outline(sheet_mv, '', set(rows), sections, out, [0])
    return out


def _takes_number(var: Any, rows: set) -> bool:
    return (
        var.id in rows
        and getattr(var, '_mirror_of', None) is None
        and not (getattr(var, '_excel_props', None) or {}).get('article')
        and row_kind(var) != 'link'
    )


def _number_outline(
    mv: Any, prefix: str, rows: set, sections: Optional[Container[str]],
    out: Dict[str, str], count: List[int],
) -> None:
    """Number ``mv``'s rows under ``prefix``, in the order ``_layout_mv``
    lays them out. ``count`` is the place last taken at this level — shared
    with the parent when ``mv`` heads no row of its own."""
    orient = _own_orient(mv)
    if orient == 'records':
        return
    if orient == 'columns':
        # Instances run across; the rows are the fields the first instance
        # names, then the container's own lines beneath the matrix.
        instances = [
            c for c in mv._components.values()
            if isinstance(c, MultiVariableBase) and not _is_presentation(c)
        ]
        if not instances:
            return
        lines = [f for inst in instances for f in inst._components.values()] + list(
            mv._components.values()
        )
        for var in lines:
            if isinstance(var, Variable) and _takes_number(var, rows):
                count[0] += 1
                out[var.id] = f"{prefix}{count[0]}"
        return
    for name in mv._component_order:
        comp = mv._components[name]
        if isinstance(comp, Variable):
            if _takes_number(comp, rows):
                count[0] += 1
                out[comp.id] = f"{prefix}{count[0]}"
            continue
        if not isinstance(comp, MultiVariableBase) or _is_presentation(comp):
            continue
        section_id = comp.id or comp.python_name or comp._name_in_parent
        if (section_id not in sections) if sections is not None else comp._is_sheet:
            continue
        own = (getattr(comp, '_excel_props', None) or {}).get('article')
        if own:
            _number_outline(comp, f"{str(own).rstrip('.')}.", rows, sections, out, [0])
        elif not comp.display_name:
            # No label, no row to number: its lines count as the parent's.
            _number_outline(comp, prefix, rows, sections, out, count)
        else:
            count[0] += 1
            number = f"{prefix}{count[0]}"
            out[section_id] = number
            _number_outline(comp, f"{number}.", rows, sections, out, [0])


def _row_number(address: str) -> int:
    """``"AB7"`` → 7 — the row an A1 ref names."""
    i = 0
    while i < len(address) and address[i].isalpha():
        i += 1
    return int(address[i:] or 0)


def _col_number(address: str) -> int:
    """``"AB7"`` → 28 — the column an A1 ref names."""
    n = 0
    for ch in address:
        if not ch.isalpha():
            break
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n


def _is_context_sheet(sheet_mv: Any) -> bool:
    """A Context sheet — lines that come from outside this model, laid out
    on this book's calendar: the container carries the ``_is_context``
    marker, or the virtual sheet standing for it does."""
    source = getattr(sheet_mv, '_sheet_source', None) or sheet_mv
    return bool(
        getattr(sheet_mv, '_is_context', False) or getattr(source, '_is_context', False)
    )


def _declares_view(node: Any) -> bool:
    """Whether a container up ``node``'s owner chain points at an
    ``ExcelView`` of its own."""
    from .view import ExcelView

    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        if isinstance(getattr(n, 'default_excel_view', None), ExcelView):
            return True
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return False


def _top_of(node: Any) -> Any:
    """The root of ``node``'s owner chain — the model a line belongs to."""
    seen: set = set()
    n, top = node, node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        top = n
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return top


def _sheet_has_constants(mv: MultiVariableBase, flatten: bool = False) -> bool:
    """Does this sheet carry any line that is NOT per-period?"""
    for comp in mv._components.values():
        if isinstance(comp, Variable):
            if not _is_periodic(comp):
                return True
        elif isinstance(comp, MultiVariableBase):
            if _is_presentation(comp):
                continue
            if _own_orient(comp) == 'columns':
                # A columns-oriented section's columns are INSTANCES. Its cells are
                # neither periods nor constants in this sense — they
                # have a geometry of their own, and counting them
                # would move a header that must not cover them. Its
                # own direct lines are rows of the sheet, though.
                if any(
                    isinstance(c, Variable) and not _is_periodic(c)
                    for c in comp._components.values()
                ):
                    return True
                continue
            if comp._is_sheet and not flatten:
                continue
            if _sheet_has_constants(comp, flatten):
                return True
    return False


def _is_periodic(var: Variable) -> bool:
    """Does this line run ALONG the time axis?

    A list (one value per period) or a tracked series does; a plain
    number does not, however many periods the model declares. The
    distinction is the whole point of a period header — it names the
    columns a series occupies, and a sheet of constants has none.
    """
    value = getattr(var, '_value', None)
    if isinstance(value, list) or type(value).__name__ == 'TrackValues':
        return True
    return bool(getattr(var, '_indexed_by', ()) or ())


def line_keys(var: Variable) -> Optional[List[str]]:
    """The keys a line runs ALONG instead of periods — a keyed list
    (``mo.Variable({'FY2023': 400, 'FY2024': 450})``) or a line on one
    finite axis that is not time (``indexed_by=[regions]``) — one per
    cell, in cell order; ``None`` for a series, a constant and a
    tracked line. Such a line stands under a caption row of these keys
    (``key_captions_by_sheet``) and counts for no period column. Public
    so a renderer captions the line by the keys the layout placed it
    under."""
    keys = getattr(var, '_keys', None)
    if keys and isinstance(getattr(var, '_value', None), list):
        return [str(k) for k in keys]
    axes: Sequence[Any] = getattr(var, '_indexed_by', None) or ()
    if len(axes) == 1 and getattr(axes[0], '_axis_kind', None) != 'time':
        labels = getattr(axes[0], '_value', None)
        if isinstance(labels, list):
            return [str(label) for label in labels]
    return None


def _sheet_has_periods(mv: MultiVariableBase, flatten: bool = False) -> bool:
    """Does this SHEET carry any per-period line?

    The gate on the period header. Without it the header rode the
    MODEL's window rather than the sheet's content, so a sheet of
    assumptions — every value a constant — printed "Jan 2026" over
    column B, and each assumption read as a January figure.

    A nested sheet's content belongs to that sheet, not this one,
    unless the caller is flattening them into one. A keyed line runs
    along its own keys, not the periods: it asks for no period header.
    """
    for comp in mv._components.values():
        if isinstance(comp, Variable):
            if _is_periodic(comp) and line_keys(comp) is None:
                return True
        elif isinstance(comp, MultiVariableBase):
            if _is_presentation(comp):
                continue
            if _own_orient(comp) == 'columns':
                # A columns-oriented section's columns are INSTANCES. Its cells are
                # neither periods nor constants in this sense — they
                # have a geometry of their own, and counting them
                # would move a header that must not cover them. Its
                # own direct lines are rows of the sheet, though.
                if any(
                    isinstance(c, Variable) and _is_periodic(c) and line_keys(c) is None
                    for c in comp._components.values()
                ):
                    return True
                continue
            if comp._is_sheet and not flatten:
                continue
            if _sheet_has_periods(comp, flatten):
                return True
    return False


def _has_captions(view: Any) -> bool:
    """Does the view name any column in words — ``meta={'unit': 'Unit'}``,
    a field's caption, the constants column's? A bare ``True`` keeps a
    column headerless, so it names nothing."""
    named = [view.meta_article_col, view.meta_unit_col, view.meta_label_col,
             view.meta_constants_col]
    named += [caption for _attr, caption in (view.meta_fields or ())]
    return any(isinstance(c, str) and c for c in named)


def _reject_rank2_emission(var: Variable) -> None:
    """Refuse to lay out a rank≥2 Variable (two or more axes).

    The one-row layout would silently paint an (A×B) buffer as A·B
    consecutive periods — plausible-looking, wrong numbers. Excel
    emission of such a Variable is not supported, so it fails loudly.

    Tracked Variables PASS: they lay out as ONE row sized by their
    time length — the display-default series row. The xlsx writer
    expands tracked lines into per-track rows BEFORE layout
    (``writer.expand_tracked_tree``); a caller that lays out the
    unexpanded tree gets the single row, holding the default series.
    """
    axes = getattr(var, '_indexed_by', ()) or ()
    if len(axes) >= 2:
        label = (getattr(var, '_display_name', None)
                 or getattr(var, '_python_name', None) or 'variable')
        raise ValueError(
            f"'{label}' is laid out along {len(axes)} axes — Excel "
            f"emission of multi-axis Variables is not supported. "
            f"Emit a slice instead (drop an axis first), or keep the "
            f"variable out of the workbook."
        )


def layout_sheet_names(sheets: Sequence[MultiVariableBase]) -> List[str]:
    """The key the layout files each sheet under, in order: its
    :func:`sheet_name_for`, suffixed ``1``, ``2``, … when an earlier sheet
    already took the name. Public so a renderer that must find a sheet
    by the layout's key keys it the same way."""
    names: List[str] = []
    taken: set = set()
    for sheet_mv in sheets:
        base = sheet_name_for(sheet_mv)
        name, suffix = base, 1
        while name in taken:
            name = f"{base}{suffix}"
            suffix += 1
        taken.add(name)
        names.append(name)
    return names


def sheet_name_for(mv: MultiVariableBase) -> str:
    """Canonical sheet name for a Sheet-roled MV.

    Single source of truth — every layer that needs to know "which
    sheet is this MV's tab?" calls this. Without one canonical
    derivation the layout / writer / translator can disagree on the
    sheet name (e.g. ``id="forecast"`` vs ``display_name="Forecast"``)
    and the translator's case-sensitive ``current_sheet`` check fails
    to suppress same-sheet qualifiers — every reference ends up
    written as ``=Sheet!Bn`` even within the same sheet.

    Priority:

      1. ``excel_layout.sheet`` — structured layout override
         (``mo.MultiVariable(excel_layout=layout)``, where ``layout``
         is any object with a ``sheet`` attribute, e.g.
         ``types.SimpleNamespace(sheet="P&L")``). Highest precedence
         because it's the explicit "rename this sheet" setting. Engine
         reads the attribute opaquely; it imports no layout type.
      2. ``sheet_name`` — explicitly set via ``excel_props={'tab': True,
         'sheet_name': '…'}``. The flag-dict path.
      3. ``display_name`` — user-facing label. ``mo.Model("forecast")``
         produces ``display_name="Forecast"`` (title-cased) which is
         the right Excel tab name.
      4. ``id`` — the qualified-path id (dotted, lowercase). Fallback
         for panel roots used by the HTML repr where neither of the
         above is set.
    """
    layout = getattr(mv, "_excel_layout", None)
    if layout is not None:
        # Read the ``sheet`` attribute off whatever object was passed;
        # don't introspect the type (the engine imports no extension types).
        layout_sheet = getattr(layout, "sheet", None)
        if layout_sheet:
            return str(layout_sheet)
    return mv.sheet_name or mv.display_name or mv.id


class LayoutEngine:
    """Walks a MultiVariable tree and assigns Excel cell addresses.

    Layout conventions:

    - Each Sheet (``_excel_props={'tab': True}``) becomes a separate Excel tab.
    - Nested non-sheet MVs become visual sections — a single header row
      (at column ``col``) followed by their components on subsequent rows.
    - Variables get one row each: column A = label, column B = reserved
      for a formula cell, columns C+ = per-period values (one cell for
      scalars; one per period for lists).
    - Row collisions are resolved by advancing to the next free row via
      :meth:`_ensure_row_free`, so two MVs can't write to the same line.

    Instance state produced by :meth:`compute_addresses`:

    - ``addresses`` — the ``{var_id: VariableAddresses}`` map consumed by
      the writer and translator.
    - ``sheet_map`` — ``{sheet_name: [var_ids_in_order]}`` for downstream
      sheet-walking code.
    - ``section_header_rows`` — ``{section_id: (row, col)}`` so custom
      emitters can place section headers at the right position.
    - ``sheets_with_key_header`` — sheets whose keyed Variables have no
      datetime axis to name columns. Informational: the keys stand in
      the caption rows below.
    - ``keyed_var_ids`` / ``key_captions_by_sheet`` — the lines that run
      along keys of their own (see :func:`line_keys`), and per sheet the
      cells of their caption rows: ``{"C3": "FY2023", …}``, each key over
      the cell that holds its value. ``key_caption_rows`` — each such
      line's caption row, ``{var_id: row}``.
    - ``unit_rows`` — ``{var_id: row}`` for a caption that heads no row
      (a records-table field, a scalar across a board) on a sheet with a
      unit column: the unit row its table reserves under the captions.
    """

    def __init__(
        self,
        roots: Sequence[LayoutRoot],
        flatten_nested_sheets: bool = False,
        totals_override: Optional[Sequence[str]] = None,
        header_rows: bool = False,
    ):
        """Store the roots; state fills in at ``compute_addresses()``.

        Args:
            roots: Top-level layout roots. Each entry is either a
                :class:`MultiVariableBase` (sheet/section subtree) or a
                :class:`Variable` (lays out as a single row on a
                synthesized sheet named after the Variable). Usually
                the Sheet/Model subtree the writer is about to emit;
                HTML repr passes one root per rendered tab.
            flatten_nested_sheets: When ``True``, sub-MVs marked
                ``_is_sheet=True`` render as sections inside their
                parent's flat layout instead of being skipped (and
                emitted as separate sheets). HTML repr's per-tab
                rendering uses this so a tab's panel reads as one
                continuous flat sheet — section headers claim real
                rows, formulas reference cells against the flat
                layout. ``to_excel`` callers leave it ``False``.
        """
        # The totals OVERRIDE: when given, replaces every sheet's
        # declared ``timeline.totals`` kinds for THIS layout — ()
        # strips totals, ('quarter',) shows quarter buckets only. None
        # keeps the declaration. A per-layout display choice, so the xlsx
        # writer never passes it.
        self._totals_override = (
            tuple(totals_override) if totals_override is not None else None
        )
        self.roots = roots
        self._flatten_nested_sheets = flatten_nested_sheets
        #: The workbook's header rows (``default_header``) — mirrors of the
        #: anchor's rows at the top of every sheet. Only a WORKBOOK lays
        #: them out; the notebook view and other renderers show the anchor
        #: as an ordinary sheet and keep their own layout.
        self._header_rows_on = header_rows
        #: ``id(sheet_mv) → SheetHeader`` for every sheet showing a header.
        self.header_by_sheet: Dict[int, Any] = {}
        #: sheet (name and ``id(mv)``) -> ``(article, unit, label)`` — the
        #: view's lead-column flags as resolved, a caption kept as text.
        #: The unit slot keeps all three of its states (a caption or
        #: ``True``, ``None``, ``False`` — see ``ResolvedView.meta_unit_col``).
        self.meta_by_sheet: Dict[Any, tuple] = {}
        self._hoisted_ids: set = set()
        #: Value addresses laid out by columns-oriented sections — their
        #: columns are INSTANCES, not periods; the timeline header
        #: width must ignore them in every renderer.
        self.board_var_ids: set = set()
        self.addresses: Dict[str, VariableAddresses] = {}
        self.sheet_map: Dict[str, list] = {}
        self.section_header_rows: Dict[str, tuple] = {}  # section_id -> (row, col)
        #: REPEATS of a section's label. A record laid out as a table row
        #: also names itself once per per-period block below the table
        #: (``orient='records'``), so one section legitimately owns
        #: several label cells. ``section_header_rows`` keeps the first —
        #: the canonical one every existing consumer reads — and the
        #: rest live here: ``section_id -> [(row, col), …]``.
        self.section_header_echoes: Dict[str, list] = {}
        #: Sections whose label shares a row with data — a record laid
        #: out as a table row, or as a row of one per-period block. The
        #: banner must NOT extend across such a row: it would style the
        #: numbers standing beside the label.
        self.inline_section_ids: set = set()
        #: Value ids whose row runs ALONG the time axis. The period
        #: header names their columns and nobody else's.
        self.period_var_ids: set = set()
        #: Value ids whose row runs along KEYS of its own (:func:`line_keys`):
        #: laid out from the first period column like a series, but off
        #: the period header's accounting — the header is the periods'.
        self.keyed_var_ids: set = set()
        #: sheet (name and ``id(mv)``) -> ``{cell: key}`` — the caption row
        #: of each keyed line, directly above it: every key over the cell
        #: that holds its value. Consecutive keyed lines with the same keys
        #: share one. Painted in the header's look by every renderer.
        self.key_captions_by_sheet: Dict[Any, Dict[str, str]] = {}
        #: var id -> the caption row each keyed line stands under — the
        #: one row the lines that share its keys share. A renderer hides
        #: it with its lines (:func:`hidden_key_caption_rows`).
        self.key_caption_rows: Dict[str, int] = {}
        #: var id -> the row its unit stands in, for a caption that heads
        #: no row of the sheet (a records-table field, a scalar across a
        #: board) on a sheet with a unit column: the UNIT ROW its table
        #: reserves directly under the captions, one unit per field
        #: column. Reserved only where a field carries a unit — like the
        #: constants column, a row that would stand empty is not taken.
        self.unit_rows: Dict[str, int] = {}
        #: sheet (name AND ``id(mv)``) -> the column periods start at.
        #: ONE number, computed where the geometry is decided, so the
        #: writer and every other renderer read it instead of
        #: re-deriving it from lead-column arithmetic and disagreeing.
        self.period_start_by_sheet: Dict[Any, int] = {}
        #: sheet -> the projected field columns declared for it,
        #: ``((attr, caption), …)`` in column order.
        self.meta_fields_by_sheet: Dict[Any, tuple] = {}
        self._meta_fields: tuple = ()
        #: Reserve a column for CONSTANTS on a sheet that also carries
        #: periods, so a rate of 12% doesn't sit under "Jan 2026" and
        #: read as a January figure.
        self._constants_col = False
        #: sheet (name and ``id(mv)``) -> the constants column's state:
        #: its caption string when the view declares one
        #: (``meta={'constants': 'Value'}``), ``True`` when the
        #: column is reserved but unnamed. Absent = no column. Without
        #: a caption the column reads as MISALIGNMENT — a run of
        #: numbers standing left of the first month under an empty
        #: header cell.
        self.constants_by_sheet: Dict[Any, Any] = {}
        #: sheet (name and ``id(mv)``) -> ``{role: column}`` for the
        #: sheet's columns: ``article``, ``label``, ``unit``, each
        #: projected field by name, ``constants`` and ``period`` (the
        #: first period column) — where each one stands, decided here
        #: once for the writer and every other renderer.
        self.columns_by_sheet: Dict[Any, Dict[str, int]] = {}
        #: sheet (name and ``id(mv)``) -> the rows above the header —
        #: 1 under a title row (``sheet={'title': True}``), else 0. The
        #: header (period captions, column captions) stands at row
        #: ``1 + top``; every renderer counts from here.
        self.top_by_sheet: Dict[Any, int] = {}
        #: sheet (name and ``id(mv)``) -> the row of its column captions,
        #: for a sheet that stands them WITHOUT a period header: a sheet
        #: of constants in a book whose sheets share their columns
        #: (``sheet={'same_columns': True}``) heads No · Line item · Unit
        #: · … · Constant like every other sheet. A sheet with a period
        #: header puts the captions in that header and is absent here.
        self.caption_rows_by_sheet: Dict[Any, int] = {}
        #: sheet (name and ``id(mv)``) -> the opening column's caption
        #: (``timeline={'opening': 'Pre-start'}``), ``True`` when
        #: headerless; absent = no opening column.
        self.opening_by_sheet: Dict[Any, Any] = {}
        self._opening_col = False
        #: sheet (name and ``id(mv)``) -> the cell the panes freeze at:
        #: the first period's column, the first row under the header and
        #: the rows it repeats. Decided here, once, for every renderer.
        self.freeze_by_sheet: Dict[Any, str] = {}
        self._freeze_facts: Dict[str, tuple] = {}
        #: line id or section id -> the number the view's
        #: ``meta['numbering']`` gives it (``'1.2'``), where it has no
        #: ``article`` of its own.
        self.articles: Dict[str, str] = {}
        #: sheet name -> its numbering scheme (``meta['numbering']``), for
        #: every sheet whose lines are numbered — so a renderer that numbers
        #: the lines itself asks :func:`number_lines` the same question the
        #: layout asked.
        self.numbering_by_sheet: Dict[str, Any] = {}
        self._numbering: Dict[str, tuple] = {}
        self.sheets_with_key_header: set = set()
        # id(sheet_mv) -> (start, grain, style) for the timeline header row.
        # Keyed by object identity, not sheet name: a model and its ``.at()``
        # clone share a display_name (so ``sheet_name_for`` collides), and
        # name-keying would let the second overwrite the first.
        self.time_headers: Dict[int, tuple] = {}
        # DEDUPED sheet name -> the same (start, grain, style) tuple — the
        # serializable projection for consumers outside this process (the
        # in-process writer keeps the id() keying above; the deduped name
        # is unique per compute_addresses pass, so no clone collision).
        self.time_headers_by_sheet: Dict[str, tuple] = {}
        #: id(sheet_mv) -> the time window its period columns follow
        #: (``None`` when the model declares none) — recorded whether or
        #: not a header shows it, because the columns follow it anyway.
        self.window_by_sheet: Dict[int, Any] = {}
        # sheet_name -> {row -> item_id that claimed it}
        self._occupied_rows: Dict[str, Dict[int, str]] = {}
        #: sheet name -> the ``id(mv)`` keys sharing it, so the
        #: period-start number is published under both spellings
        #: consumers look it up by.
        self._sheet_ids_by_name: Dict[str, list] = {}
        #: sheet -> declared subtotal interleave (view + window), and
        #: the built plans — ``totals.TotalsPlan`` under BOTH the name
        #: and ``id(mv)`` keys, like every other per-sheet product.
        self._totals_decl: Dict[str, tuple] = {}
        self._totals_ids: Dict[str, int] = {}
        self.totals_by_sheet: Dict[Any, Any] = {}
        #: Track column groups (``tracks='compare'``): per sheet, the
        #: group's ``(stride, words, hidden_slots)`` — how many columns
        #: a period takes, what each subcolumn is called, and which of
        #: them the writer hides (a blend operand left out of the
        #: selection but needed by the splice formula). ONE number,
        #: computed where the geometry is decided — the writer, the totals and
        #: the blend links all read it here instead of re-deriving it.
        self.groups_by_sheet: Dict[Any, tuple] = {}
        #: group head id() → its assigned row, so slot copies land
        #: beside their head instead of on rows of their own.
        self._group_rows: Dict[tuple, int] = {}
        self._group_stride: int = 1
        #: var id → its slot, for the passes that run after layout
        #: (``_apply_totals`` moves cells by slot) and hold only ids.
        self._slot_by_id: Dict[str, int] = {}
        #: sheet → {row: outline level} — native Excel ROW groups
        #: following the section tree, so a section collapses in the
        #: workbook itself. Built by accumulation: every enclosing
        #: section bumps its span by one as the recursion unwinds, so
        #: depth needs no bookkeeping.
        self.row_outline_by_sheet: Dict[str, Dict[int, int]] = {}
        #: id(sheet_mv) → this layout's name for it — the key of every
        #: name-keyed map above. A workbook tab may be spelled otherwise
        #: (Excel-safe, case-insensitively unique), so a writer finds a
        #: sheet's plans through this, never through its tab title.
        self.sheet_name_by_id: Dict[int, str] = {}

    def _ensure_row_free(self, sheet_name: str, row: int) -> int:
        """Return the first unoccupied row >= *row* for the given sheet."""
        occupied = self._occupied_rows.get(sheet_name)
        if not occupied:
            return row
        while row in occupied:
            row += 1
        return row

    def _mark_row(self, sheet_name: str, row: int, item_id: str) -> None:
        """Record that *row* on *sheet_name* is occupied by *item_id*."""
        self._occupied_rows.setdefault(sheet_name, {})[row] = item_id

    def compute_addresses(self) -> Dict[str, VariableAddresses]:
        """Walk all sheet-level MVs and assign addresses to every Variable."""
        sheets = self._collect_sheets()

        from ...core.time import (
            DEFAULT_TIME_LABEL_STYLE,
            TIME_LABEL_STYLES,
            DefaultWindow,
            resolve_default_window,
        )
        from .view import resolve_excel_view

        # The book's own calendar and view — the first sheet that is not a
        # Context and resolves a window — for a Context sheet to stand on:
        # lines from outside this model, drawn as a tab of THIS book.
        book_window = None
        book_root = None
        for sheet_mv in sheets:
            if _is_context_sheet(sheet_mv):
                continue
            anchor = next(self._iter_sheet_variables(sheet_mv), None)
            w = resolve_default_window(anchor) if anchor is not None else None
            if w is not None and w.grain is not None and w.start is not None:
                book_window, book_root = w, _top_of(anchor)
                break

        for sheet_mv, sheet_name in zip(sheets, layout_sheet_names(sheets)):
            # Unique per-sheet key for internal bookkeeping (``_occupied_rows``,
            # ``sheet_map``). ``sheet_name_for`` collides for a model and its
            # ``.at()`` clone (shared display_name); without a distinct key their
            # row spaces merge and the second tab's rows slide down the sheet.
            self.sheet_map[sheet_name] = []
            self.sheet_name_by_id[id(sheet_mv)] = sheet_name

            needs_key_header = self._sheet_needs_key_header(sheet_mv)
            if needs_key_header:
                self.sheets_with_key_header.add(sheet_name)

            # Resolve the view + time window from a Variable inside the sheet:
            # variables keep their _owner→model link even when the sheet MV is a
            # detached virtual wrapper (the same trick the orient lookup uses).
            anchor_var = next(self._iter_sheet_variables(sheet_mv), None)
            # ``resolve_excel_view`` returns the global house-default snapshot
            # for ``None``, so no separate fallback is needed.
            view = resolve_excel_view(anchor_var)
            # A Context sheet wears the book's view (label style, lead
            # columns, the header's presence) unless it declares one of its
            # own — its dates must read as the model's tabs read them.
            is_context = _is_context_sheet(sheet_mv)
            if is_context and book_root is not None and not _declares_view(anchor_var):
                view = resolve_excel_view(book_root)
            # Dense-nesting rule for THIS sheet's section walk: blank
            # rows between sections come from the view (None = the
            # classic single gap).
            self._section_gap = (
                view.nesting_gap_rows
                if view.nesting_gap_rows is not None else 1
            )
            # Lead metadata columns (article number / unit): labels shift
            # right by the article column, values by the unit column too.
            # Kept RAW, not coerced to bool: a string value is the
            # column's caption (``meta={'unit': 'Unit'}``) and the
            # writer needs it downstream. Every use here is
            # truthiness-only, and a non-empty caption is truthy — so
            # the geometry is unchanged either way. The unit keeps its
            # unset ``None`` apart from ``False`` too: a renderer tells
            # them apart (``units_hidden``) where it writes the text.
            self._meta_article = view.meta_article_col or False
            self._meta_unit = view.meta_unit_col
            #: Caption only — the label column exists regardless.
            self._meta_label = view.meta_label_col or False
            #: Projected field columns — one per ``meta['fields']``
            #: entry, between the unit column and the values.
            self._meta_fields = tuple(view.meta_fields or ())
            self.meta_fields_by_sheet[sheet_name] = self._meta_fields
            self.meta_fields_by_sheet[id(sheet_mv)] = self._meta_fields
            window = (
                resolve_default_window(anchor_var) if anchor_var is not None else None
            )
            if window is None and is_context and book_window is not None:
                # The book's start and grain only — never its periods: the
                # header's length is the lines' own extent, and a window key
                # distinct from the model's keeps the sheet's captions text,
                # never formulas off the model's header rows. The lines
                # themselves keep no time of their own; only the layout
                # stands on the calendar.
                window = DefaultWindow(book_window.start, book_window.grain, None)
            self.window_by_sheet[id(sheet_mv)] = window

            # Transpose (``orient='down'``) is supported for a FLAT sheet only —
            # all direct children are Variables, no sub-sections. Anything nested
            # falls back to the default across-layout (byte-identical).
            sheet_vars = [
                c for c in sheet_mv._components.values() if isinstance(c, Variable)
            ]
            content = [
                c for c in sheet_mv._components.values() if not _is_presentation(c)
            ]
            is_flat = len(sheet_vars) == len(content)
            is_transposed = (view.orient or "across") == "down" and is_flat
            _hdr_mirrors = self._plan_header(sheet_mv, window, is_transposed, view)
            if is_transposed:
                # A transposed sheet runs periods DOWN — its layout
                # never reserves lead columns, so the writer must not
                # shift either. The unit column goes; the view's word
                # on whether units show at all stays.
                self._meta_article = self._meta_label = False
                self._meta_unit = None if self._meta_unit else self._meta_unit
            # Keyed by BOTH the raw sheet name (external consumers use
            # layout names verbatim) and the MV's identity (the writer
            # re-sanitizes titles — a name key would silently miss).
            _meta_triple = (
                self._meta_article, self._meta_unit, self._meta_label
            )
            self.meta_by_sheet[sheet_name] = _meta_triple
            self.meta_by_sheet[id(sheet_mv)] = _meta_triple

            _extra_header_rows = 0
            # Timeline header: a period-label row at row 1 when the sheet has a
            # resolved window (start + grain) and the view doesn't suppress it.
            # Across-layout only — a transposed sheet runs periods DOWN, so its
            # variable-name header_row already plays that role.
            want_time_header = (
                not is_transposed
                # A columns-oriented section runs INSTANCES across, not
                # periods — a period header over it would lie.
                and _own_orient(sheet_mv) != 'columns'
                # Content, not the model's window, decides: a sheet of
                # constants has no columns for a period header to name.
                # (Subsumes the records case — a table of scalar
                # records is exactly that sheet.) A declared header puts
                # periods on every sheet.
                and (_sheet_has_periods(sheet_mv, self._flatten_nested_sheets)
                     or bool(_hdr_mirrors))
                and window is not None
                and window.grain is not None
                and window.start is not None
                and view.time_header is not False
            )
            if want_time_header:
                assert window is not None  # narrowed by want_time_header
                style = view.time_label_format or DEFAULT_TIME_LABEL_STYLE
                _tot_kinds = (
                    self._totals_override
                    if self._totals_override is not None
                    else tuple(view.time_totals or ())
                )
                # No buckets on a Context sheet: a line that is not this
                # model's to re-grain has no rule to sum or carry by.
                if _tot_kinds and not is_context:
                    # Subtotal columns ride the time header — resolved
                    # here where the view and window are in hand,
                    # applied after addresses exist (_apply_totals).
                    self._totals_decl[sheet_name] = (
                        tuple(_tot_kinds),
                        window.start,
                        window.grain,
                        style,
                        view.time_totals_word or "Total",
                    )
                    self._totals_ids[sheet_name] = id(sheet_mv)
                    from .totals import _coarser as _tot_coarser

                    _usable = [
                        k for k in _tot_kinds
                        if _tot_coarser(k, window.grain)
                    ]
                    # The hierarchical header: one row per bucket kind
                    # ABOVE the month row — data moves down with it.
                    _extra_header_rows = len(_usable)
                if style not in TIME_LABEL_STYLES:
                    raise ValueError(
                        f"ExcelView.time_label_format must be one of "
                        f"{TIME_LABEL_STYLES}; got {style!r}."
                    )
                self.time_headers[id(sheet_mv)] = (window.start, window.grain, style)
                self.time_headers_by_sheet[sheet_name] = (
                    window.start,
                    window.grain,
                    style,
                )

            # Track column groups: the sheet's STRIDE is the widest
            # slot any of its lines carries — resolved here, once,
            # because the address arithmetic below multiplies by it.
            # Slot words and hidden slots ride along for the writer.
            self._group_stride = 1
            _words: Dict[int, str] = {}
            _hidden: set = set()
            for _v in self._iter_sheet_variables(sheet_mv):
                _slot = getattr(_v, '_col_slot', None)
                if _slot is None:
                    continue
                self._group_stride = max(self._group_stride, _slot + 1)
                _w = getattr(_v, '_col_word', None)
                if _w and _slot not in _words:
                    _words[_slot] = _w
                if getattr(_v, '_col_hidden', False):
                    _hidden.add(_slot)
            if self._group_stride > 1:
                _group = (
                    self._group_stride,
                    tuple(_words.get(i, "") for i in range(self._group_stride)),
                    tuple(sorted(_hidden)),
                )
                self.groups_by_sheet[sheet_name] = _group
                self.groups_by_sheet[id(sheet_mv)] = _group
                if want_time_header:
                    # The track-word row under the period labels —
                    # data moves down with it.
                    _extra_header_rows += 1

            # A sheet with no period header still heads its columns
            # where the book's sheets share them: the captions stand in
            # the header row, the same words in the same places as on
            # every other sheet. Without ``same_columns`` such a sheet
            # keeps row 1 for data, as it always did.
            want_captions = (
                not want_time_header
                and not is_transposed
                and _own_orient(sheet_mv) not in ('columns', 'records')
                and bool(view.sheet_same_columns)
                and _has_captions(view)
            )

            # A constants column only where it earns its place: the
            # sheet has periods AND something that isn't periodic.
            # A records-oriented sheet never gets one: its scalars live
            # in the fields TABLE, and a constants column would stand
            # as dead space between the table and the first month.
            # A constant among the header rows (a status at the top of
            # every sheet) is one of the sheet's own; ``same_columns``
            # keeps the column on every sheet regardless, so a period
            # stands in one column across the book. On a captioned sheet
            # of constants the column its values stand in IS that
            # column: it takes its caption and its place, and moves no
            # cell (only a periodic line steps past it).
            self._constants_col = (
                (want_time_header or want_captions)
                and _own_orient(sheet_mv) != 'records'
                and (
                    bool(view.sheet_same_columns)
                    or _sheet_has_constants(
                        sheet_mv, self._flatten_nested_sheets
                    )
                    # A scalar row among the header mirrors (a date, a
                    # status) needs a column of its own - but only a
                    # view that names the constants column or shares
                    # the columns gets it: a model that asks for
                    # nothing new keeps its book, cell for cell.
                    or (
                        bool(view.meta_constants_col or view.sheet_same_columns)
                        and any(not _is_periodic(m) for _s, m in _hdr_mirrors)
                    )
                )
            )
            if self._constants_col:
                _const_caption = view.meta_constants_col or True
                self.constants_by_sheet[sheet_name] = _const_caption
                self.constants_by_sheet[id(sheet_mv)] = _const_caption

            # The opening column: the point just before the first
            # period, on a sheet whose columns ARE periods.
            self._opening_col = bool(
                want_time_header
                and _own_orient(sheet_mv) != 'records'
                and view.time_opening
            )
            if self._opening_col:
                self.opening_by_sheet[sheet_name] = view.time_opening
                self.opening_by_sheet[id(sheet_mv)] = view.time_opening
            if not is_transposed:
                _roles = self._lead_columns()
                self.columns_by_sheet[sheet_name] = _roles
                self.columns_by_sheet[id(sheet_mv)] = _roles
            # A title row above everything: the header moves down with it.
            _top = 1 if view.sheet_title else 0
            self.top_by_sheet[sheet_name] = _top
            self.top_by_sheet[id(sheet_mv)] = _top
            if want_captions:
                self.caption_rows_by_sheet[sheet_name] = _top + 1
                self.caption_rows_by_sheet[id(sheet_mv)] = _top + 1

            # A keyed line names its columns in the caption row above it,
            # so a sheet of keyed lines reserves no key header row.
            start_row = _top + (
                2 + _extra_header_rows
                if (want_time_header or want_captions) else 1
            )

            if _hdr_mirrors and want_time_header:
                start_row = self._layout_header_rows(
                    _hdr_mirrors, start_row, 2 if self._meta_article else 1, sheet_name
                )
            if is_transposed:
                self._layout_mv_transposed(
                    sheet_mv, header_row=start_row, start_col=1, sheet_name=sheet_name
                )
            else:
                self._layout_mv(
                    sheet_mv, row=start_row,
                    col=2 if self._meta_article else 1,
                    sheet_name=sheet_name,
                )
            if sheet_name in self.time_headers_by_sheet:
                self._sheet_ids_by_name.setdefault(sheet_name, []).append(
                    id(sheet_mv)
                )
            if sheet_name in self.key_captions_by_sheet:
                self.key_captions_by_sheet[id(sheet_mv)] = (
                    self.key_captions_by_sheet[sheet_name]
                )
            if view.meta_numbering and not is_transposed:
                self._numbering[sheet_name] = (sheet_mv, view.meta_numbering)
            self._freeze_facts[sheet_name] = (
                id(sheet_mv), bool(self._meta_article), bool(self._meta_unit), _top,
                bool(want_time_header), list(_hdr_mirrors), want_captions,
            )

        self._record_period_starts()
        self._apply_totals()
        self._record_freezes()
        self._record_numbers()
        return self.addresses

    def _record_numbers(self) -> None:
        """Number each numbered sheet's sections and lines as an outline, in
        the order they stand (``meta={'numbering': 'section'}``)."""
        for name, (sheet_mv, scheme) in self._numbering.items():
            self.numbering_by_sheet[name] = scheme
            self.articles.update(number_lines(
                sheet_mv,
                [v for v in self.sheet_map.get(name, [])
                 if getattr(self.addresses.get(v), 'name', None)],
                scheme, self.section_header_rows,
            ))

    def _record_freezes(self) -> None:
        """Where each sheet's panes freeze: left of the first period (the
        label and lead columns — and the opening column — stay in view),
        and under the header, the rows it repeats below it included."""
        for name, (oid, article, unit, top, headed, mirrors,
                   captioned) in self._freeze_facts.items():
            col = self.period_start_by_sheet.get(oid) or 2 + int(article) + int(unit)
            plan = self.totals_by_sheet.get(name)
            depth = (plan.depth if plan is not None else 1) + (
                1 if name in self.groups_by_sheet else 0
            )
            row = 1 + top + (depth if headed else 1 if captioned else 0)
            if mirrors and headed:
                row = max(
                    _row_number(self.addresses[m.id].values[0]) for _s, m in mirrors
                ) + 1
            cell = f"{self._col_to_letter(col)}{row}"
            self.freeze_by_sheet[name] = cell
            self.freeze_by_sheet[oid] = cell
    
    def _lead_columns(self) -> Dict[str, int]:
        """The current sheet's lead columns, by role — the same
        arithmetic :meth:`_create_var_address` places cells by."""
        label = 2 if self._meta_article else 1
        roles: Dict[str, int] = {'label': label}
        if self._meta_article:
            roles['article'] = 1
        col = label + 1
        if self._meta_unit:
            roles['unit'] = col
            col += 1
        for attr, _caption in self._meta_fields:
            roles.setdefault(attr, col)
            col += 1
        if self._constants_col:
            roles['constants'] = col
            col += 1
        if self._opening_col:
            roles['opening'] = col
        return roles

    def _plan_header(self, sheet_mv: MultiVariableBase, window: Any,
                     is_transposed: bool, view: Any = None) -> list:
        """This sheet's header rows (see :mod:`.header`): the rows it
        shows mirrored at its top. Rows that live on the sheet itself
        (the anchor's own sheet) are its content, not mirrors — unless
        the view repeats them there too (``timeline={'repeat_on_anchor':
        True}``), when the sheet shows them at its top as well."""
        if (not self._header_rows_on or is_transposed
                or _own_orient(sheet_mv) in ('columns', 'records')
                or window is None or window.start is None):
            return []
        from ...core.time import resolve_default_window
        from .header import (SheetHeader, header_rows, mirror_of,
                             resolve_default_header, window_key)
        source = getattr(sheet_mv, '_sheet_source', None) or sheet_mv
        rows = header_rows(resolve_default_header(source))
        # A header row counts periods of ITS window: over a sheet of
        # another window (a quarterly block under a monthly header) its
        # columns would stand under the wrong periods.
        key = window_key(window)
        rows = [r for r in rows if window_key(resolve_default_window(r)) == key]
        if not rows:
            return []
        home_ids = {id(v) for v in self._iter_sheet_variables(sheet_mv)}
        repeat = bool(getattr(view, 'time_repeat_on_anchor', None))
        plan = SheetHeader(key=key)
        for row in rows:
            if id(row) in home_ids:
                plan.home.append(row)
                if not repeat:
                    continue
            plan.mirrors.append((row, mirror_of(row)))
        self.header_by_sheet[id(sheet_mv)] = plan
        return plan.mirrors

    def _layout_header_rows(self, mirrors: list, row: int, col: int,
                            sheet_name: str) -> int:
        """Place the header mirrors at the top of the sheet, then the
        section gap; returns the first row for the sheet's content."""
        for _src, mirror in mirrors:
            row = self._ensure_row_free(sheet_name, row)
            self._mark_row(sheet_name, row, mirror.id)
            self.addresses[mirror.id] = self._create_var_address(mirror, row, col)
            self.sheet_map.setdefault(sheet_name, []).append(mirror.id)
            row += 1
        return row + self._section_gap

    def _sheet_needs_key_header(self, sheet_mv: MultiVariableBase) -> bool:
        """Return True when a sheet has Variables with labeled keys
        but no datetime-valued Variable.

        The top-row header prints the key labels so columns are readable;
        when a date/period axis Variable is present, its labels play
        that role and a second header is redundant.
        """
        has_keys = False
        has_date_axis = False
        for var in self._iter_sheet_variables(sheet_mv):
            if var.var_type == 'datetime':
                has_date_axis = True
            if getattr(var, '_keys', None):
                has_keys = True
        return has_keys and not has_date_axis

    def _iter_sheet_variables(self, sheet_mv: MultiVariableBase) -> Iterator[Variable]:
        """Yield every Variable reachable within a sheet, descending through
        any depth of nested non-sheet sections. By default, nested sheets
        are pruned — they become their own tabs and their contents are
        yielded when those sheets are processed. With
        ``flatten_nested_sheets=True``, all sub-MVs (including
        ``_is_sheet=True`` ones) are walked, since they render as
        sections in the parent's flat layout."""
        stack: List[MultiVariableBase] = [sheet_mv]
        while stack:
            mv = stack.pop()
            for name in mv._component_order:
                comp = mv._components[name]
                if isinstance(comp, Variable):
                    yield comp
                elif isinstance(comp, MultiVariableBase):
                    if _is_presentation(comp):
                        continue
                    if comp._is_sheet and not self._flatten_nested_sheets:
                        continue
                    stack.append(comp)

    def _collect_sheets(self) -> List[MultiVariableBase]:
        """Collect the sheet-level layout targets from the roots.

        Default mode (``flatten_nested_sheets=False``): MV roots and
        every ``_is_sheet=True`` descendant become sheets — preserves
        Excel's auto-promotion behaviour for ``to_excel``. Variable
        roots are wrapped via :meth:`_make_variable_sheet`.

        Flat-tab mode (``flatten_nested_sheets=True``): only the
        explicit roots become sheets. Nested ``_is_sheet=True``
        descendants stay inside their parent's tree as sections —
        :meth:`_layout_mv` no longer skips them in this mode.
        """
        sheets: List[MultiVariableBase] = []
        for root in self.roots:
            if isinstance(root, Variable):
                sheets.append(self._make_variable_sheet(root))
                continue
            if isinstance(root, MultiVariableBase):
                if root not in sheets:
                    sheets.append(root)
                if not self._flatten_nested_sheets:
                    for _name, comp in root.walk():
                        if (
                            isinstance(comp, MultiVariableBase)
                            and comp._is_sheet
                            and not _is_presentation(comp)
                        ):
                            if comp not in sheets:
                                sheets.append(comp)
        return sheets

    @staticmethod
    def _make_variable_sheet(var: Variable) -> MultiVariableBase:
        """Wrap a single Variable as a layout-able pseudo-sheet.

        Used when callers (HTML repr in particular) want to lay out
        a single Variable as if it were a one-row sheet — same
        column conventions, same address shape — without first
        having to attach it under an MV. The sheet inherits the
        Variable's ``display_name`` (humanized from ``python_name``
        when no explicit label is set) so the tab reads as the
        Variable itself.
        """
        from ...core.multi_variable import MultiVariable

        # An anonymous Variable (no explicit label, no python_name)
        # would otherwise expose its synthesized auto-id — prefer a
        # generic "Variable" tab name in that case.
        if var._display_name:
            tab_name = var._display_name
        elif var.python_name:
            tab_name = var.display_name
        else:
            tab_name = 'Variable'

        sheet = object.__new__(MultiVariable)
        sheet._role = 'sheet'
        sheet._display_name = tab_name
        sheet._components = {var.id: var}
        sheet._component_order = [var.id]
        sheet._parent = None
        sheet._name_in_parent = None
        sheet._qualified_id = None
        sheet._python_name = None
        sheet._excel_props = {'tab': True}
        # The wrapper must not LOSE the hidden flag: a hidden Variable
        # wrapped into a synthetic tab stays a hidden sheet.
        if effective_hidden(var):
            sheet._excel_props['hidden'] = True
        return sheet
    
    def _layout_mv(self, mv: MultiVariableBase, row: int, col: int, sheet_name: str) -> int:
        """Recursively lay out an MV and its children.

        Returns the next available row after this MV's contents.
        """
        # Columns orientation (``ExcelView(orient='columns')``):
        # REPEATING children spread ACROSS — one column per sub-MV
        # (instance), one row per scalar field. The matrix form for
        # instances: metrics × instances instead of N stacked
        # blocks. List fields don't fit a scalar matrix and are
        # skipped.
        if _own_orient(mv) == 'columns':
            return self._layout_mv_columns(mv, row, col, sheet_name)
        # Records orientation: one ROW per child record, one COLUMN
        # per field — a table. The transpose of ``'columns'`` above,
        # and the shape a LIST of records actually has; the default
        # stacked layout turns N records into N sections, each
        # repeating the same fields.
        if _own_orient(mv) == 'records':
            return self._layout_mv_records(mv, row, col, sheet_name)
        # Whether the previous laid-out child was a plain Variable row —
        # a section header that follows one needs the same breathing gap
        # it leaves behind itself, or it glues to the row above.
        prev_was_var = False
        # The keys and caption row of the keyed line laid out last, while
        # the next line follows it directly: lines that share the keys
        # share the caption row.
        prev_keyed: Optional[tuple] = None
        for name in mv._component_order:
            comp = mv._components[name]

            if isinstance(comp, MultiVariableBase):
                if _is_presentation(comp):
                    continue
                if comp._is_sheet and not self._flatten_nested_sheets:
                    continue
                prev_keyed = None
                # Use the fully-qualified id so same-named sections on different
                # sheets (e.g. control.dates vs timing.flags.dates) don't collide.
                section_id = comp.id or comp.python_name or comp._name_in_parent

                explicit_row = comp._excel_props.get("row")
                explicit_col = comp._excel_props.get("col")
                if prev_was_var and explicit_row is None:
                    # Gap BEFORE a section that follows variable rows —
                    # symmetric to the gap every section leaves after
                    # itself (the after-gap is added at the branch tail).
                    row += getattr(self, '_section_gap', 1)
                prev_was_var = False
                proposed = explicit_row if explicit_row is not None else row
                sec_row = max(proposed, row)
                sec_row = self._ensure_row_free(sheet_name, sec_row)
                sec_col = explicit_col if explicit_col is not None else col

                self._mark_row(sheet_name, sec_row, section_id)
                self.section_header_rows[section_id] = (sec_row, sec_col)

                # Header-row total (``excel_props={'header_row': True}``
                # on a child Variable): its VALUES print on the section
                # header's own row — the financial-statement form where
                # "2.1 · Operating expenses" carries the subtotal and the
                # components sit under it. The label cell stays the
                # section's; the child gets value addresses only.
                for _hname in comp._component_order:
                    _hchild = comp._components[_hname]
                    if (isinstance(_hchild, Variable)
                            and (_hchild._excel_props or {}).get('header_row')):
                        _haddr = self._create_var_address(
                            _hchild, sec_row, sec_col
                        )
                        _haddr = VariableAddresses(
                            name='', formula=_haddr.formula,
                            values=_haddr.values,
                        )
                        self.addresses[_hchild.id] = _haddr
                        self.sheet_map.setdefault(sheet_name, []).append(
                            _hchild.id
                        )
                        # Exactly ONE row rides the header; siblings
                        # (and misplaced header_row rows anywhere else)
                        # lay out as ordinary rows — never dropped.
                        self._hoisted_ids.add(id(_hchild))
                        break

                child_start = self._ensure_row_free(sheet_name, sec_row + 1)
                next_row = self._layout_mv(comp, child_start, sec_col, sheet_name)
                # ROW group: this section's whole span sits one level
                # deeper than wherever the section itself stands. The
                # header row stays OUTSIDE the group — it is the
                # summary Excel puts the ± on (summaryBelow=False).
                _lvl = self.row_outline_by_sheet.setdefault(sheet_name, {})
                for _r in range(sec_row + 1, next_row):
                    _lvl[_r] = min(7, _lvl.get(_r, 0) + 1)

                row = max(row, next_row + getattr(self, '_section_gap', 1))
            elif isinstance(comp, Variable):
                if id(comp) in self._hoisted_ids:
                    continue  # rides its parent's header row (above)
                head = getattr(comp, '_col_head', None)
                if head is not None:
                    # A track-slot copy: beside its head, not below it.
                    # The head laid out first (expansion order) and
                    # left its row here; the copy takes the same row
                    # at its own slot's columns, label cell NONE — the
                    # row already has its name.
                    head_row = self._group_rows.get(
                        (sheet_name, id(head))
                    )
                    if head_row is not None:
                        addr = self._create_var_address(
                            comp, head_row, col
                        )
                        self.addresses[comp.id] = VariableAddresses(
                            name='', formula=addr.formula,
                            values=addr.values,
                        )
                        self._slot_by_id[comp.id] = (
                            getattr(comp, '_col_slot', 0) or 0
                        )
                        self.sheet_map.setdefault(
                            sheet_name, []
                        ).append(comp.id)
                        continue
                keys = line_keys(comp)
                if keys is None:
                    prev_keyed = None
                elif prev_keyed is None or prev_keyed[0] != keys:
                    # A keyed line stands under a caption row of its keys,
                    # directly above it: both rows free, the caption's
                    # first.
                    cap_row = self._ensure_row_free(sheet_name, row)
                    while cap_row + 1 in self._occupied_rows.get(sheet_name, {}):
                        cap_row = self._ensure_row_free(sheet_name, cap_row + 1)
                    self._mark_row(sheet_name, cap_row, f"{comp.id}.__keys__")
                    prev_keyed = (keys, cap_row)
                    row = cap_row + 1
                row = self._ensure_row_free(sheet_name, row)
                key = comp.id
                self._mark_row(sheet_name, row, key)
                addr = self._create_var_address(comp, row, col)
                self.addresses[key] = addr
                self.sheet_map.setdefault(sheet_name, []).append(key)
                if keys is not None and prev_keyed is not None:
                    self.key_caption_rows[key] = prev_keyed[1]
                    captions = self.key_captions_by_sheet.setdefault(sheet_name, {})
                    for cell, label in zip(addr.values, keys):
                        captions[
                            f"{self._col_to_letter(_col_number(cell))}{prev_keyed[1]}"
                        ] = label
                if getattr(comp, '_col_slot', None) is not None:
                    self._group_rows[(sheet_name, id(comp))] = row
                    self._slot_by_id[key] = comp._col_slot or 0
                row += 1
                prev_was_var = True

        return row
    
    def _layout_mv_records(
        self, mv: MultiVariableBase, sec_row: int, sec_col: int,
        sheet_name: str,
    ) -> int:
        """Table layout: one row per child record.

        The mirror image of :meth:`_layout_mv_columns`. A container of
        like records is a LIST, and the stacked default turns N of them
        into N sections of the same fields — the shape nobody writes by
        hand in Excel.

        A record's fields split by whether they carry a period axis, and
        each half gets the form it actually has:

        * SCALAR fields become a table — fields across, records down.
          Field names ride the header row as the first record's label
          cells (the trick the columns layout uses one axis over), so
          every renderer paints them through the ordinary label path.
        * PER-PERIOD fields cannot share that table: its columns are
          fields, and a series' columns are periods. Each such field
          gets its own block below — title, then one row per record,
          values landing in the sheet's own timeline columns. That is
          the pivot every analyst builds by hand: "Revenue" once, the
          records under it.

        A record therefore names itself once in the table and once per
        block; ``section_header_echoes`` carries the repeats. The
        container's own direct lines — a total under the table — follow
        beneath as ordinary rows.
        """
        records = [
            (name, c) for name, c in mv._components.items()
            if isinstance(c, MultiVariableBase) and not _is_presentation(c)
        ]
        lines = [
            c for c in mv._components.values()
            if isinstance(c, Variable) and id(c) not in self._hoisted_ids
        ]
        if not records and not lines:
            return sec_row

        def _is_series(f: Variable) -> bool:
            v = getattr(f, '_value', None)
            return isinstance(v, list) or type(v).__name__ == 'TrackValues'

        # Field order is first-appearance across all records, so a
        # record that omits a field leaves a gap rather than a shift.
        scalar_order: list[str] = []
        series_order: list[str] = []
        for _, rec in records:
            for fname, f in rec._components.items():
                if not isinstance(f, Variable):
                    continue
                bucket = series_order if _is_series(f) else scalar_order
                if fname not in bucket and fname not in (
                    scalar_order if bucket is series_order else series_order
                ):
                    bucket.append(fname)

        # Which record names each column: the FIRST that carries the
        # field. Fixed up front rather than raced for during the walk —
        # records of one class usually agree on a field's display name,
        # and when they don't the header must still be stable instead
        # of "whoever was written last".
        header_owner = {
            fname: next(
                id(rec) for _n, rec in records
                if isinstance(rec._components.get(fname), Variable)
            )
            for fname in scalar_order
        }

        row = sec_row
        if scalar_order:
            header_row = self._ensure_row_free(sheet_name, row)
            self._mark_row(sheet_name, header_row, f"{mv.id}.__header__")
            row = header_row + 1
            # The unit row: a field's caption heads no row of the sheet,
            # so its unit stands under the caption, in the row the table
            # reserves for every field's — where the sheet has a unit
            # column and a field has a unit to put there. The record
            # that names a column (``header_owner``) names its unit too:
            # a unit only a later record carries does not relabel the
            # column, and reserves no row on its own.
            unit_row = self._unit_row([
                rec._components[fname] for _n, rec in records
                for fname in scalar_order
                if header_owner.get(fname) == id(rec)
            ], row, sheet_name, f"{mv.id}.__units__")
            if unit_row is not None:
                row = unit_row + 1
            for i, (rname, rec) in enumerate(records):
                row = self._ensure_row_free(sheet_name, row)
                rec_id = rec.id or rec.python_name or rname
                self._mark_row(sheet_name, row, rec_id)
                # The row's identity cell, painted by the section-header
                # pass.
                self.section_header_rows[rec_id] = (row, sec_col)
                self.inline_section_ids.add(rec_id)
                for j, fname in enumerate(scalar_order):
                    f = rec._components.get(fname)
                    if not isinstance(f, Variable):
                        continue
                    # Past the lead metadata columns, exactly as
                    # ``_create_var_address`` steps past them: the unit
                    # column names a ROW's unit and a record row has no
                    # single one, but the column is spoken for on this
                    # sheet — writing a field into it would land the
                    # value under "Unit".
                    col_j = sec_col + 1 + j + (
                        1 if getattr(self, '_meta_unit', False) else 0
                    )
                    cell = f"{self._col_to_letter(col_j)}{row}"
                    # Only the first record carrying this field owns the
                    # header cell — the rest are value-only, exactly as
                    # the columns layout does for its sibling instances.
                    name_addr = (
                        f"{self._col_to_letter(col_j)}{header_row}"
                        if header_owner.get(fname) == id(rec)
                        else ''
                    )
                    self.addresses[f.id] = VariableAddresses(
                        name=name_addr, formula=cell, values=[cell],
                    )
                    self.board_var_ids.add(f.id)
                    self.sheet_map.setdefault(sheet_name, []).append(f.id)
                    if name_addr and unit_row is not None:
                        self.unit_rows[f.id] = unit_row
                row += 1

        # Period blocks start PAST the fields table. The table's columns
        # are fields and a series' columns are months — sharing the
        # column space would paint a field's caption under "Feb 2026":
        # the month masthead (read back from these value addresses)
        # would stretch over the field columns. The
        # shift keeps the two column vocabularies in disjoint ranges;
        # the masthead follows the addresses on its own.
        series_col = sec_col + len(scalar_order)

        for fname in series_order:
            owners = [
                (rname, rec, rec._components[fname])
                for rname, rec in records
                if isinstance(rec._components.get(fname), Variable)
            ]
            if not owners:
                continue
            title_row = self._ensure_row_free(sheet_name, row + (1 if row > sec_row else 0))
            self._mark_row(sheet_name, title_row, f"{mv.id}.{fname}.__block__")
            row = title_row + 1
            for k, (rname, rec, f) in enumerate(owners):
                row = self._ensure_row_free(sheet_name, row)
                self._mark_row(sheet_name, row, f.id)
                addr = self._create_var_address(f, row, series_col)
                # The block is titled once, by its first record's field —
                # the label cell the writer already knows how to paint.
                # Every row's own label is the RECORD's name, which the
                # section-header pass writes from the echo below.
                self.addresses[f.id] = VariableAddresses(
                    name=(
                        f"{self._col_to_letter(sec_col)}{title_row}"
                        if k == 0 else ''
                    ),
                    formula=addr.formula,
                    values=addr.values,
                )
                self.sheet_map.setdefault(sheet_name, []).append(f.id)
                rec_id = rec.id or rec.python_name or rname
                self.inline_section_ids.add(rec_id)
                if rec_id in self.section_header_rows:
                    self.section_header_echoes.setdefault(
                        rec_id, []
                    ).append((row, sec_col))
                else:
                    self.section_header_rows[rec_id] = (row, sec_col)
                row += 1

        # The container's own DIRECT lines (a total under the table)
        # follow beneath the table and the blocks — never dropped. A
        # series stands in the columns the blocks use, so the month
        # masthead stays over them; a scalar where the sheet's scalars
        # stand. The label is the line's own, in the label column.
        for var in lines:
            row = self._ensure_row_free(sheet_name, row)
            self._mark_row(sheet_name, row, var.id)
            addr = self._create_var_address(
                var, row, series_col if _is_periodic(var) else sec_col
            )
            self.addresses[var.id] = VariableAddresses(
                name=f"{self._col_to_letter(sec_col)}{row}",
                formula=addr.formula, values=addr.values, opening=addr.opening,
            )
            self.sheet_map.setdefault(sheet_name, []).append(var.id)
            row += 1

        return row

    def _layout_mv_columns(
        self, mv: MultiVariableBase, sec_row: int, sec_col: int,
        sheet_name: str,
    ) -> int:
        """Matrix layout for a columns-oriented section: instance
        display names across ``sec_row`` (one column each, from the
        first period column), scalar fields as rows beneath — the first
        instance's field owns the label cell in the label column;
        sibling instances' same-name fields get value-only addresses in
        their column."""
        instances = [
            (name, c) for name, c in mv._components.items()
            if isinstance(c, MultiVariableBase) and not _is_presentation(c)
        ]
        # The board's own DIRECT lines: a scalar (a total across the
        # instances) and a series alike — never dropped. A line hoisted
        # onto the parent's header row already stands there.
        lines = [
            (name, c) for name, c in mv._components.items()
            if isinstance(c, Variable) and id(c) not in self._hoisted_ids
        ]
        scalars = [
            (name, c) for name, c in lines
            if not isinstance(getattr(c, '_value', None), list)
        ]
        series = [
            (name, c) for name, c in lines
            if isinstance(getattr(c, '_value', None), list)
        ]

        def _line_row(var: Variable, row: int) -> int:
            # A direct line is a row of the sheet: a scalar's value
            # where the sheet's scalars stand, a series' cells in the
            # period columns. The series is no board cell — the
            # masthead names its columns and stretches over them.
            row = self._ensure_row_free(sheet_name, row)
            self._mark_row(sheet_name, row, var.id)
            self.addresses[var.id] = self._create_var_address(var, row, sec_col)
            if not _is_periodic(var):
                self.board_var_ids.add(var.id)
            self.sheet_map.setdefault(sheet_name, []).append(var.id)
            return row + 1

        # The board's columns are instances, run across the way periods
        # are: its first column is the first period's — past the label
        # and every lead column the sheet reserves (unit, fields,
        # constants, opening), the steps ``_create_var_address`` takes
        # for a series. Started at the label's neighbour, the board sat
        # under the unit column, and the unit text of its first field
        # row was written over the first instance's value.
        first = sec_col + 1 + (
            1 if getattr(self, '_meta_unit', False) else 0
        ) + len(getattr(self, '_meta_fields', ()) or ()) + (
            1 if getattr(self, '_constants_col', False) else 0
        ) + (1 if getattr(self, '_opening_col', False) else 0)
        row = sec_row
        if instances:
            # Instance headers ride the section-header machinery: one
            # position per instance on the section's own row.
            for j, (iname, inst) in enumerate(instances):
                inst_id = inst.id or inst.python_name or iname
                self.section_header_rows[inst_id] = (row, first + j)
            field_order: list[str] = []
            for _, inst in instances:
                for fname, f in inst._components.items():
                    if (isinstance(f, Variable)
                            and not isinstance(getattr(f, '_value', None), list)
                            and fname not in field_order):
                        field_order.append(fname)
            row += 1
            for fname in field_order:
                row = self._ensure_row_free(sheet_name, row)
                self._mark_row(sheet_name, row, f"{mv.id}.{fname}")
                for j, (_, inst) in enumerate(instances):
                    f = inst._components.get(fname)
                    if not isinstance(f, Variable):
                        continue
                    cell = f"{self._col_to_letter(first + j)}{row}"
                    name_addr = (
                        f"{self._col_to_letter(sec_col)}{row}" if j == 0 else ''
                    )
                    self.addresses[f.id] = VariableAddresses(
                        name=name_addr, formula=cell, values=[cell],
                    )
                    self.board_var_ids.add(f.id)
                    self.sheet_map.setdefault(sheet_name, []).append(f.id)
                row += 1
            # The section's own direct lines follow beneath the matrix,
            # in declaration order.
            for _name, var in lines:
                row = _line_row(var, row)
            return row
        if scalars:
            # Scalar children: names across one row, their units in the
            # row beneath where the sheet has a unit column (a caption
            # across heads no row to put one beside), values beneath.
            hdr_row = row + 1
            self._mark_row(sheet_name, hdr_row, f"{mv.id}.__hdr__")
            unit_row = self._unit_row(
                [var for _n, var in scalars], hdr_row + 1, sheet_name, f"{mv.id}.__units__",
            )
            val_row = (hdr_row if unit_row is None else unit_row) + 1
            for j, (name, var) in enumerate(scalars):
                cell = f"{self._col_to_letter(first + j)}{val_row}"
                label = f"{self._col_to_letter(first + j)}{hdr_row}"
                self.addresses[var.id] = VariableAddresses(
                    name=label, formula=cell, values=[cell],
                )
                self.board_var_ids.add(var.id)
                self.sheet_map.setdefault(sheet_name, []).append(var.id)
                if unit_row is not None:
                    self.unit_rows[var.id] = unit_row
            self._mark_row(sheet_name, val_row, f"{mv.id}.__val__")
            row = val_row + 1
            # A series has no place across; it follows beneath as a row.
            for _name, var in series:
                row = _line_row(var, row)
            return row
        if series:
            for _name, var in series:
                row = _line_row(var, row)
            return row
        return row + 1

    def _unit_row(
        self, lines: Sequence[Variable], row: int, sheet_name: str, item_id: str,
    ) -> Optional[int]:
        """Reserve the unit row of a table or board whose captions head no
        row — the first free row at ``row`` — when the sheet has a unit
        column and any of ``lines`` carries a unit; ``None`` otherwise."""
        if not getattr(self, '_meta_unit', False) or not any(
            getattr(var, '_unit', None) is not None for var in lines
        ):
            return None
        unit_row = self._ensure_row_free(sheet_name, row)
        self._mark_row(sheet_name, unit_row, item_id)
        return unit_row

    def _record_period_starts(self) -> None:
        """Fix each sheet's period-start column from the laid-out cells.

        Read back rather than recomputed: the lead columns, the
        constants column and the records layouts each move the first
        value column, and one authority beats three agreeing
        arithmetics.
        """
        for sheet, var_ids in self.sheet_map.items():
            cols = [
                _col_number(self.addresses[v].values[0])
                for v in var_ids
                if v in self.period_var_ids
                and v not in self.board_var_ids
                and v in self.addresses
                and self.addresses[v].values
            ]
            if not cols:
                continue
            start = min(cols)
            self.period_start_by_sheet[sheet] = start
            for oid in self._sheet_ids_by_name.get(sheet, []):
                self.period_start_by_sheet[oid] = start
            roles = self.columns_by_sheet.get(sheet)
            if roles is not None:
                roles['period'] = start

    def _apply_totals(self) -> None:
        """Interleave subtotal columns into every declaring sheet.

        Runs AFTER addresses and period starts exist: native month
        cells move right past the inserted bucket columns — the
        ADDRESSES move, so the translator and every renderer follow by
        construction — and each periodic line gains one cell per
        bucket, recorded on the plan for the writer and any other
        consumer.
        """
        from .totals import build_totals_plan

        for sheet, decl in self._totals_decl.items():
            kinds, start, grain, style, totals_word = decl
            base = self.period_start_by_sheet.get(sheet)
            if not base:
                continue
            vids = [
                v for v in self.sheet_map.get(sheet, [])
                if v in self.period_var_ids
                and v not in self.board_var_ids
                and v in self.addresses
            ]
            n = max(
                (len(self.addresses[v].values) for v in vids), default=0
            )
            stride = self.groups_by_sheet.get(sheet, (1,))[0]
            plan = build_totals_plan(kinds, grain, start, n, base, stride)
            if plan is None:
                continue
            plan.style = style
            plan.totals_word = totals_word
            for v in vids:
                a = self.addresses[v]
                row = _row_number(a.values[0])
                slot = self._slot_by_id.get(v, 0)
                moved = [
                    f"{self._col_to_letter(plan.cell(('month', i), slot))}{row}"
                    for i in range(len(a.values))
                ]
                self.addresses[v] = VariableAddresses(
                    name=a.name, formula=moved[0], values=moved,
                    opening=a.opening,
                )
            self.totals_by_sheet[sheet] = plan
            oid = self._totals_ids.get(sheet)
            if oid is not None:
                self.totals_by_sheet[oid] = plan

    def _create_var_address(self, var: Variable, row: int, start_col: int) -> VariableAddresses:
        """Create cell addresses for a single Variable.

        Column layout: ``start_col`` (typically A) holds the label;
        ``start_col + 1`` onwards (B, C, …) hold values. ``formula``
        points at the first value cell and is currently equivalent to
        ``values[0]`` — kept as a separate field so consumers can
        distinguish "the cell that holds a formula" from "the first
        cell of a per-period values run."
        """
        _reject_rank2_emission(var)
        name_addr = f"{self._col_to_letter(start_col)}{row}"
        values_start = start_col + 1 + (
            1 if getattr(self, '_meta_unit', False) else 0
        ) + len(getattr(self, '_meta_fields', ()) or ())
        periodic = _is_periodic(var)
        # A keyed line takes a series' place — its cells start where
        # the periods do — but it is no series: it counts for no period
        # column, and its cells run one per key, whatever the stride.
        keyed = periodic and line_keys(var) is not None
        opening_addr = None
        if periodic:
            (self.keyed_var_ids if keyed else self.period_var_ids).add(var.id)
            # Step past the constants column. A scalar keeps the first
            # value column — it is not a January figure and must not
            # be printed as one.
            if getattr(self, '_constants_col', False):
                values_start += 1
            # …and past the opening column, which the row owns too.
            if getattr(self, '_opening_col', False):
                opening_addr = f"{self._col_to_letter(values_start)}{row}"
                values_start += 1

        # Track column groups: a period is ``stride`` columns wide and
        # this line's cells sit at its SLOT inside each group. Stride 1
        # / slot 0 — every workbook without ``tracks='compare'`` —
        # reduces to one column per period.
        stride = 1 if keyed else getattr(self, '_group_stride', 1)
        slot = 0 if keyed else getattr(var, '_col_slot', 0) or 0
        formula_addr = (
            f"{self._col_to_letter(values_start + slot)}{row}"
            if periodic else f"{self._col_to_letter(values_start)}{row}"
        )

        value = var._value
        if type(value).__name__ == 'TrackValues':
            # One row sized by the track time length — the display-
            # default series (the writer path expands tracked lines
            # BEFORE layout; this branch serves callers that lay out
            # the unexpanded tree).
            n = value.time_length or 1
            value_addrs = [
                f"{self._col_to_letter(values_start + i * stride + slot)}{row}"
                for i in range(n)
            ]
        elif isinstance(value, list) and len(value) > 1:
            value_addrs = [
                f"{self._col_to_letter(values_start + i * stride + slot)}{row}"
                for i in range(len(value))
            ]
        else:
            value_addrs = [f"{self._col_to_letter(values_start)}{row}"]

        return VariableAddresses(
            name=name_addr,
            formula=formula_addr,
            values=value_addrs,
            opening=opening_addr,
        )

    def _layout_mv_transposed(
        self, mv: MultiVariableBase, header_row: int, start_col: int, sheet_name: str
    ) -> None:
        """Lay out a flat sheet TRANSPOSED — periods down, Variables across.

        Each Variable becomes a column: its label sits in ``header_row`` and its
        values run down beneath it. Formulas and recurrences follow for free
        because they reference the (now vertical) value addresses, not a fixed
        across-pattern.
        """
        col = start_col
        for name in mv._component_order:
            comp = mv._components[name]
            if isinstance(comp, Variable):
                self.addresses[comp.id] = self._create_var_address_down(
                    comp, header_row, col
                )
                self.sheet_map.setdefault(sheet_name, []).append(comp.id)
                col += 1

    def _create_var_address_down(
        self, var: Variable, header_row: int, col: int
    ) -> VariableAddresses:
        """Addresses for a Variable laid out as a COLUMN (label on top, values
        running down) — the transpose of :meth:`_create_var_address`."""
        _reject_rank2_emission(var)
        name_addr = f"{self._col_to_letter(col)}{header_row}"
        values_start_row = header_row + 1
        value = var._value
        n = len(value) if isinstance(value, list) and len(value) > 1 else 1
        value_addrs = [
            f"{self._col_to_letter(col)}{values_start_row + i}" for i in range(n)
        ]
        return VariableAddresses(
            name=name_addr,
            formula=value_addrs[0],
            values=value_addrs,
        )

    @staticmethod
    def _col_to_letter(col: int) -> str:
        """Convert a 1-based column number to Excel letters.

        ``1 → "A"``, ``26 → "Z"``, ``27 → "AA"``, ``52 → "AZ"``,
        ``53 → "BA"``, etc. (Excel's base-26 with no zero digit.)
        """
        result = ""
        while col > 0:
            col -= 1
            result = chr(65 + col % 26) + result
            col //= 26
        return result


