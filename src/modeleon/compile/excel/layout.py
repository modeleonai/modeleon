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

from typing import Any, Dict, Iterator, List, Optional, Sequence, Union

from .addresses import VariableAddresses
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable


def _is_presentation(comp: object) -> bool:
    """A nested ExcelView is a named view recipe — presentation
    metadata that layout must never emit as rows or sections. (The
    ``default_excel_view`` slot bypasses adoption entirely; this guard
    covers the ADOPTED named views.)"""
    from .view import ExcelView

    return isinstance(comp, ExcelView)

LayoutRoot = Union[Variable, MultiVariableBase]


def effective_hidden(node: object) -> bool:
    """True когда узел или ЛЮБОЙ предок несёт ``excel_props={'hidden': True}``.

    Скрытая строка — не пропущенная: она размечается, считается и несёт
    формулы/подписи (оси чартов выводятся из адресов), но writer прячет
    её нативной скрытой строкой, грид не рисует, граф фильтрует. Каскад
    ``_owner`` → ``_parent`` — тот же, что у ``resolve_excel_view``.
    """
    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        if (getattr(n, '_excel_props', None) or {}).get('hidden'):
            return True
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return False


def _own_orient(mv: Any) -> Optional[str]:
    """The layout orientation ``mv`` DECLARES for itself.

    Read from the MV's own ``ExcelView`` — the one place a printed form
    is described, so a collection that wants to print as a table says
    so the same way it says anything else about its form. The legacy
    ``excel_props={'orient': …}`` spelling still wins where it is set,
    so existing models are byte-identical.

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
                # A board's columns are INSTANCES. Its cells are
                # neither periods nor constants in this sense — they
                # have a geometry of their own, and counting them
                # would move a header that must not cover them.
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


def _sheet_has_periods(mv: MultiVariableBase, flatten: bool = False) -> bool:
    """Does this SHEET carry any per-period line?

    The gate on the period header. Without it the header rode the
    MODEL's window rather than the sheet's content, so «Должности» —
    a register of salary bands, every value a constant — printed
    «Jan 2026» over column B and its 500 read as a January figure.

    A nested sheet's content belongs to that sheet, not this one,
    unless the caller is flattening them into one.
    """
    for comp in mv._components.values():
        if isinstance(comp, Variable):
            if _is_periodic(comp):
                return True
        elif isinstance(comp, MultiVariableBase):
            if _is_presentation(comp):
                continue
            if _own_orient(comp) == 'columns':
                # A board's columns are INSTANCES. Its cells are
                # neither periods nor constants in this sense — they
                # have a geometry of their own, and counting them
                # would move a header that must not cover them.
                continue
            if comp._is_sheet and not flatten:
                continue
            if _sheet_has_periods(comp, flatten):
                return True
    return False


def _reject_rank2_emission(var: Variable) -> None:
    """Refuse to lay out a rank≥2 (multi-axis) Variable (§14.4 gate).

    The one-row layout would silently paint an (A×B) buffer as A·B
    consecutive periods — plausible-looking, wrong numbers. A loud
    error until axis-aware emission (column groups) lands.

    Tracked Variables PASS (P3): they lay out as ONE row sized by
    their time length — the display-default series row. The xlsx
    writer expands tracked lines into per-track rows BEFORE layout
    (``writer.expand_tracked_tree``); the grid-address path keys the
    single row by the variable's qpath and reads the default series
    from ``VariableRuntime.values``.
    """
    axes = getattr(var, '_indexed_by', ()) or ()
    if len(axes) >= 2:
        label = (getattr(var, '_display_name', None)
                 or getattr(var, '_python_name', None) or 'variable')
        raise ValueError(
            f"'{label}' is laid out along {len(axes)} axes — Excel "
            f"emission of multi-axis Variables is not supported yet. "
            f"Emit a slice instead (drop an axis first), or keep the "
            f"variable out of the workbook."
        )


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
         (``mo.MultiVariable(excel_layout=ExcelLayout(sheet="P&L"))``).
         Highest precedence because it's the explicit "rename this
         sheet" tag in the layout system. Engine reads the attribute
         opaquely; the layout type lives in pro.
      2. ``sheet_name`` — explicitly set via ``excel_props={'tab': True,
         'sheet_name': '…'}``. The legacy flag-dict path.
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
        # don't introspect the type (engine stays free of pro imports).
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
    - ``sheets_with_key_header`` — sheets that need a top-row key header
      (dict-keyed Variables without a datetime axis to name columns).
    """

    def __init__(
        self,
        roots: Sequence[LayoutRoot],
        flatten_nested_sheets: bool = False,
        totals_override: Optional[Sequence[str]] = None,
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
        # The VIEW LENS: when given, replaces every sheet's declared
        # ``timeline.totals`` kinds for THIS layout — () strips totals,
        # ('quarter',) shows quarter buckets only. None keeps the
        # declaration. A viewing choice, so the writer's export path
        # never passes it.
        self._totals_override = (
            tuple(totals_override) if totals_override is not None else None
        )
        self.roots = roots
        self._flatten_nested_sheets = flatten_nested_sheets
        self.meta_by_sheet: Dict[Any, tuple] = {}
        self._hoisted_ids: set = set()
        #: Value addresses laid out by columns-oriented boards — their
        #: columns are INSTANCES, not periods; the timeline header
        #: width must ignore them on both surfaces.
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
        #: sheet (name AND ``id(mv)``) -> the column periods start at.
        #: ONE number, computed where the geometry is decided, so the
        #: writer, the run-state and the grid stop re-deriving it from
        #: lead-column arithmetic and stop disagreeing about it.
        self.period_start_by_sheet: Dict[Any, int] = {}
        #: sheet -> the projected field columns declared for it,
        #: ``((attr, caption), …)`` in column order.
        self.meta_fields_by_sheet: Dict[Any, tuple] = {}
        self._meta_fields: tuple = ()
        #: Reserve a column for CONSTANTS on a sheet that also carries
        #: periods, so a rate of 12% doesn't sit under «Jan 2026» and
        #: read as a January figure.
        self._constants_col = False
        #: sheet (name and ``id(mv)``) -> the constants column's state:
        #: its caption string when the view declares one
        #: (``meta={'constants': 'Значение'}``), ``True`` when the
        #: column is reserved but unnamed. Absent = no column. Without
        #: a caption the column reads as MISALIGNMENT — a run of
        #: numbers standing left of the first month under an empty
        #: header cell — which is exactly how it was reported.
        self.constants_by_sheet: Dict[Any, Any] = {}
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
        # sheet_name -> {row -> item_id that claimed it}
        self._occupied_rows: Dict[str, Dict[int, str]] = {}
        #: sheet name -> the ``id(mv)`` keys sharing it, so the
        #: period-start number is published under both spellings the
        #: writer and the run-state read.
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
        #: them the writer hides (a blend operand that was unchecked
        #: but must exist for the splice formula). ONE number, computed
        #: where the geometry is decided — the writer, the totals and
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
        #: mirroring the section tree, so «свернуть секцию» works in
        #: the downloaded file the way it does in the grid. Built by
        #: accumulation: every enclosing section bumps its span by
        #: one as the recursion unwinds, so depth needs no bookkeeping.
        self.row_outline_by_sheet: Dict[str, Dict[int, int]] = {}

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
            resolve_default_window,
        )
        from .view import resolve_excel_view

        for sheet_mv in sheets:
            # Unique per-sheet key for internal bookkeeping (``_occupied_rows``,
            # ``sheet_map``). ``sheet_name_for`` collides for a model and its
            # ``.at()`` clone (shared display_name); without a distinct key their
            # row spaces merge and the second tab's rows slide down the sheet.
            base_name = sheet_name_for(sheet_mv)
            sheet_name = base_name
            _suffix = 1
            while sheet_name in self.sheet_map:
                sheet_name = f"{base_name}{_suffix}"
                _suffix += 1
            self.sheet_map[sheet_name] = []

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
            # Dense-nesting law for THIS sheet's section walk: blank
            # rows between sections come from the view (None = the
            # classic single gap).
            self._section_gap = (
                view.nesting_gap_rows
                if view.nesting_gap_rows is not None else 1
            )
            # Lead metadata columns (№ / ед.изм.): labels shift right
            # by the article column, values by the unit column too.
            # Kept RAW, not coerced to bool: a string value is the
            # column's caption (``meta={'unit': 'Ед. изм.'}``) and the
            # writer needs it downstream. Every use here is
            # truthiness-only, and a non-empty caption is truthy — so
            # the geometry is unchanged either way.
            self._meta_article = view.meta_article_col or False
            self._meta_unit = view.meta_unit_col or False
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
            if is_transposed:
                # A transposed sheet runs periods DOWN — its layout
                # never reserves lead columns, so the writer must not
                # shift either.
                self._meta_article = self._meta_unit = False
                self._meta_label = False
            # Keyed by BOTH the raw sheet name (the grid/run path uses
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
                # A columns-oriented board runs INSTANCES across, not
                # periods — a period header over it would lie.
                and _own_orient(sheet_mv) != 'columns'
                # Content, not the model's window, decides: a sheet of
                # constants has no columns for a period header to name.
                # (Subsumes the records case — a register of scalar
                # records is exactly that sheet.)
                and _sheet_has_periods(sheet_mv, self._flatten_nested_sheets)
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
                if _tot_kinds:
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
                    # data moves down with it, exactly the grid's own
                    # rowShift.
                    _extra_header_rows += 1

            # A constants column only where it earns its place: the
            # sheet has periods AND something that isn't periodic.
            # A records-oriented sheet never gets one: its scalars live
            # in the fields TABLE, and a constants column would stand
            # as dead space between the table and the first month.
            self._constants_col = (
                want_time_header
                and _own_orient(sheet_mv) != 'records'
                and _sheet_has_constants(
                    sheet_mv, self._flatten_nested_sheets
                )
            )
            if self._constants_col:
                _const_caption = view.meta_constants_col or True
                self.constants_by_sheet[sheet_name] = _const_caption
                self.constants_by_sheet[id(sheet_mv)] = _const_caption

            start_row = (
                2 + _extra_header_rows
                if (needs_key_header or want_time_header) else 1
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

        self._record_period_starts()
        self._apply_totals()
        return self.addresses
    
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
        # Обёртка не должна ТЕРЯТЬ служебность: скрытая переменная,
        # завёрнутая в синтетический таб, остаётся скрытым листом.
        if effective_hidden(var):
            sheet._excel_props['hidden'] = True
        return sheet
    
    def _layout_mv(self, mv: MultiVariableBase, row: int, col: int, sheet_name: str) -> int:
        """Recursively lay out an MV and its children.

        Returns the next available row after this MV's contents.
        """
        # Columns orientation (``ExcelView(orient='columns')``):
        # REPEATING children spread ACROSS — one column per sub-MV
        # (instance), one row per scalar field. The board form for
        # instances: показатели × проекты instead of N stacked
        # blocks. List fields don't fit a scalar matrix and are
        # skipped.
        if _own_orient(mv) == 'columns':
            return self._layout_mv_columns(mv, row, col, sheet_name)
        # Records orientation: one ROW per child record, one COLUMN
        # per field — a table. The transpose of ``'columns'`` above,
        # and the shape a LIST of records actually has; the default
        # stacked layout turns 29 records into 29 sections and ~350
        # rows of the same five fields repeating.
        if _own_orient(mv) == 'records':
            return self._layout_mv_records(mv, row, col, sheet_name)
        for name in mv._component_order:
            comp = mv._components[name]
            
            if isinstance(comp, MultiVariableBase):
                if _is_presentation(comp):
                    continue
                if comp._is_sheet and not self._flatten_nested_sheets:
                    continue
                # Use the fully-qualified id so same-named sections on different
                # sheets (e.g. control.dates vs timing.flags.dates) don't collide.
                section_id = comp.id or comp.python_name or comp._name_in_parent

                explicit_row = comp._excel_props.get("row")
                explicit_col = comp._excel_props.get("col")
                proposed = explicit_row if explicit_row is not None else row
                sec_row = max(proposed, row)
                sec_row = self._ensure_row_free(sheet_name, sec_row)
                sec_col = explicit_col if explicit_col is not None else col

                self._mark_row(sheet_name, sec_row, section_id)
                self.section_header_rows[section_id] = (sec_row, sec_col)

                # Header-row total (``excel_props={'header_row': True}``
                # on a child Variable): its VALUES print on the section
                # header's own row — the financial-book form where
                # «1.2 · Взносы» carries the subtotal and the
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
                row = self._ensure_row_free(sheet_name, row)
                key = comp.id
                self._mark_row(sheet_name, row, key)
                addr = self._create_var_address(comp, row, col)
                self.addresses[key] = addr
                self.sheet_map.setdefault(sheet_name, []).append(key)
                if getattr(comp, '_col_slot', None) is not None:
                    self._group_rows[(sheet_name, id(comp))] = row
                    self._slot_by_id[key] = comp._col_slot or 0
                row += 1
        
        return row
    
    def _layout_mv_records(
        self, mv: MultiVariableBase, sec_row: int, sec_col: int,
        sheet_name: str,
    ) -> int:
        """Table layout: one row per child record.

        The mirror image of :meth:`_layout_mv_columns`. A collection is
        a LIST of like records, and the stacked default turns 29 of them
        into 29 sections of the same five fields — the shape nobody
        writes by hand in Excel.

        A record's fields split by whether they carry a period axis, and
        each half gets the form it actually has:

        * SCALAR fields become a table — fields across, records down.
          Field names ride the header row as the first record's label
          cells (the trick the columns layout uses one axis over), so
          both renderers paint them through the ordinary label path.
        * PER-PERIOD fields cannot share that table: its columns are
          fields, and a series' columns are periods. Each such field
          gets its own block below — title, then one row per record,
          values landing in the sheet's own timeline columns. That is
          the pivot every analyst builds by hand: «Выручка» once, the
          projects under it.

        A record therefore names itself once in the table and once per
        block; ``section_header_echoes`` carries the repeats.
        """
        records = [
            (name, c) for name, c in mv._components.items()
            if isinstance(c, MultiVariableBase) and not _is_presentation(c)
        ]
        if not records:
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
        if not scalar_order and not series_order:
            return sec_row

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
            for i, (rname, rec) in enumerate(records):
                row = self._ensure_row_free(sheet_name, row)
                rec_id = rec.id or rec.python_name or rname
                self._mark_row(sheet_name, row, rec_id)
                # The row's identity cell, painted by the section-header
                # pass on both surfaces.
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
                    # value under «Ед. изм.».
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
                row += 1

        # Period blocks start PAST the fields table. The table's columns
        # are fields and a series' columns are months — sharing the
        # column space painted «Город» under «Feb 2026» and «Сумма»
        # under «Aug 2026»: the month masthead (read back from these
        # value addresses) stretched over the passport columns. The
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

        return row

    def _layout_mv_columns(
        self, mv: MultiVariableBase, sec_row: int, sec_col: int,
        sheet_name: str,
    ) -> int:
        """Matrix layout for a columns-oriented section: instance
        display names across ``sec_row`` (one column each, from B),
        scalar fields as rows beneath — the first instance's field
        owns the label cell in column A; sibling instances' same-name
        fields get value-only addresses in their column."""
        instances = [
            (name, c) for name, c in mv._components.items()
            if isinstance(c, MultiVariableBase) and not _is_presentation(c)
        ]
        scalars = [
            (name, c) for name, c in mv._components.items()
            if isinstance(c, Variable) and not isinstance(
                getattr(c, '_value', None), list
            )
        ]
        row = sec_row
        if instances:
            # Instance headers ride the section-header machinery: one
            # position per instance on the section's own row.
            for j, (iname, inst) in enumerate(instances):
                inst_id = inst.id or inst.python_name or iname
                self.section_header_rows[inst_id] = (row, sec_col + 1 + j)
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
                    cell = f"{self._col_to_letter(sec_col + 1 + j)}{row}"
                    name_addr = (
                        f"{self._col_to_letter(sec_col)}{row}" if j == 0 else ''
                    )
                    self.addresses[f.id] = VariableAddresses(
                        name=name_addr, formula=cell, values=[cell],
                    )
                    self.board_var_ids.add(f.id)
                    self.sheet_map.setdefault(sheet_name, []).append(f.id)
                row += 1
            # The board's own DIRECT scalar rows (a Total across the
            # instances) follow beneath the matrix — never dropped.
            for name, var in scalars:
                row = self._ensure_row_free(sheet_name, row)
                self._mark_row(sheet_name, row, var.id)
                cell = f"{self._col_to_letter(sec_col + 1)}{row}"
                self.addresses[var.id] = VariableAddresses(
                    name=f"{self._col_to_letter(sec_col)}{row}",
                    formula=cell, values=[cell],
                )
                self.board_var_ids.add(var.id)
                self.sheet_map.setdefault(sheet_name, []).append(var.id)
                row += 1
            return row
        if scalars:
            # Scalar children: names across one row, values beneath.
            hdr_row = row + 1
            val_row = row + 2
            for j, (name, var) in enumerate(scalars):
                cell = f"{self._col_to_letter(sec_col + 1 + j)}{val_row}"
                label = f"{self._col_to_letter(sec_col + 1 + j)}{hdr_row}"
                self.addresses[var.id] = VariableAddresses(
                    name=label, formula=cell, values=[cell],
                )
                self.board_var_ids.add(var.id)
                self.sheet_map.setdefault(sheet_name, []).append(var.id)
            self._mark_row(sheet_name, hdr_row, f"{mv.id}.__hdr__")
            self._mark_row(sheet_name, val_row, f"{mv.id}.__val__")
            return val_row + 1
        return row + 1

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

    def _apply_totals(self) -> None:
        """Interleave subtotal columns into every declaring sheet.

        Runs AFTER addresses and period starts exist: native month
        cells move right past the inserted bucket columns — the
        ADDRESSES move, so the translator and both surfaces follow by
        construction — and each periodic line gains one cell per
        bucket, recorded on the plan for the writer and the wire.
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
        if periodic:
            self.period_var_ids.add(var.id)
            # Step past the constants column. A scalar keeps the first
            # value column — it is not a January figure and must not
            # be printed as one.
            if getattr(self, '_constants_col', False):
                values_start += 1
        formula_addr = (
            f"{self._col_to_letter(values_start + (getattr(var, '_col_slot', 0) or 0))}{row}"
            if periodic else f"{self._col_to_letter(values_start)}{row}"
        )

        # Track column groups: a period is ``stride`` columns wide and
        # this line's cells sit at its SLOT inside each group. Stride 1
        # / slot 0 — every book without a compare lens — reduces to the
        # historical arithmetic byte for byte.
        stride = getattr(self, '_group_stride', 1)
        slot = getattr(var, '_col_slot', 0) or 0

        value = var._value
        if type(value).__name__ == 'TrackValues':
            # One row sized by the track time length — the display-
            # default series (the writer path expands tracked lines
            # BEFORE layout; this branch serves the grid-address path).
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


