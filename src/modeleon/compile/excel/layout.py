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

from typing import Dict, Iterator, List, Sequence, Union

from .addresses import VariableAddresses
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable

LayoutRoot = Union[Variable, MultiVariableBase]


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
        self.roots = roots
        self._flatten_nested_sheets = flatten_nested_sheets
        self.addresses: Dict[str, VariableAddresses] = {}
        self.sheet_map: Dict[str, list] = {}
        self.section_header_rows: Dict[str, tuple] = {}  # section_id -> (row, col)
        self.sheets_with_key_header: set = set()
        # sheet_name -> {row -> item_id that claimed it}
        self._occupied_rows: Dict[str, Dict[int, str]] = {}

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
        
        for sheet_mv in sheets:
            # Prefer the user-visible label over the qualified id —
            # non-sheet MVs (used as panel roots in HTML repr) have
            # ``sheet_name`` = None and a dotted ``id`` like ``t.h``.
            # Falling back to ``display_name`` keeps cell addresses
            # readable (``Acme!B2``, not ``t.h!B2``).
            sheet_name = (
                sheet_mv.sheet_name
                or sheet_mv.display_name
                or sheet_mv.id
            )
            self.sheet_map[sheet_name] = []
            needs_header = self._sheet_needs_key_header(sheet_mv)
            if needs_header:
                self.sheets_with_key_header.add(sheet_name)
            start_row = 2 if needs_header else 1
            self._layout_mv(sheet_mv, row=start_row, col=1, sheet_name=sheet_name)
        
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
                        if isinstance(comp, MultiVariableBase) and comp._is_sheet:
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
        return sheet
    
    def _layout_mv(self, mv: MultiVariableBase, row: int, col: int, sheet_name: str) -> int:
        """Recursively lay out an MV and its children.

        Returns the next available row after this MV's contents.
        """
        for name in mv._component_order:
            comp = mv._components[name]
            
            if isinstance(comp, MultiVariableBase):
                if comp._is_sheet and not self._flatten_nested_sheets:
                    continue
                section_id = comp.python_name or comp._name_in_parent or comp.id

                explicit_row = comp._excel_props.get("row")
                explicit_col = comp._excel_props.get("col")
                proposed = explicit_row if explicit_row is not None else row
                sec_row = max(proposed, row)
                sec_row = self._ensure_row_free(sheet_name, sec_row)
                sec_col = explicit_col if explicit_col is not None else col

                self._mark_row(sheet_name, sec_row, section_id)
                self.section_header_rows[section_id] = (sec_row, sec_col)

                child_start = self._ensure_row_free(sheet_name, sec_row + 1)
                next_row = self._layout_mv(comp, child_start, sec_col, sheet_name)

                row = max(row, next_row + 1)
            elif isinstance(comp, Variable):
                row = self._ensure_row_free(sheet_name, row)
                key = comp.id
                self._mark_row(sheet_name, row, key)
                addr = self._create_var_address(comp, row, col)
                self.addresses[key] = addr
                self.sheet_map.setdefault(sheet_name, []).append(key)
                row += 1
        
        return row
    
    def _create_var_address(self, var: Variable, row: int, start_col: int) -> VariableAddresses:
        """Create cell addresses for a single Variable.

        Column layout: ``start_col`` (typically A) holds the label;
        ``start_col + 1`` onwards (B, C, …) hold values. ``formula``
        points at the first value cell and is currently equivalent to
        ``values[0]`` — kept as a separate field so consumers can
        distinguish "the cell that holds a formula" from "the first
        cell of a per-period values run."
        """
        name_addr = f"{self._col_to_letter(start_col)}{row}"
        values_start = start_col + 1
        formula_addr = f"{self._col_to_letter(values_start)}{row}"

        value = var._value
        if isinstance(value, list) and len(value) > 1:
            value_addrs = [
                f"{self._col_to_letter(values_start + i)}{row}"
                for i in range(len(value))
            ]
        else:
            value_addrs = [f"{self._col_to_letter(values_start)}{row}"]

        return VariableAddresses(
            name=name_addr,
            formula=formula_addr,
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


