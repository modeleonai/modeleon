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

from typing import Dict, Iterator, List

from .addresses import VariableAddresses
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable


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

    def __init__(self, roots: List[MultiVariableBase]):
        """Store the roots; state fills in at ``compute_addresses()``.

        Args:
            roots: Top-level MultiVariableBase instances. Usually the
                Sheet/Model subtree the writer is about to emit.
        """
        self.roots = roots
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
            sheet_name = sheet_mv.sheet_name or sheet_mv.id
            self.sheet_map[sheet_name] = []
            needs_header = self._sheet_needs_key_header(sheet_mv)
            if needs_header:
                self.sheets_with_key_header.add(sheet_name)
            start_row = 2 if needs_header else 1
            self._layout_mv(sheet_mv, row=start_row, col=1, sheet_name=sheet_name)
        
        return self.addresses
    
    @staticmethod
    def _sheet_needs_key_header(sheet_mv: MultiVariableBase) -> bool:
        """Return True when a sheet has Variables with labeled keys
        but no datetime-valued Variable.

        The top-row header prints the key labels so columns are readable;
        when a date/period axis Variable is present, its labels play
        that role and a second header is redundant.
        """
        has_keys = False
        has_date_axis = False
        for var in LayoutEngine._iter_sheet_variables(sheet_mv):
            if var.var_type == 'datetime':
                has_date_axis = True
            if getattr(var, '_keys', None):
                has_keys = True
        return has_keys and not has_date_axis

    @staticmethod
    def _iter_sheet_variables(sheet_mv: MultiVariableBase) -> Iterator[Variable]:
        """Yield every Variable reachable within a sheet, descending through
        any depth of nested non-sheet sections. Nested sheets are pruned —
        they become their own tabs and their contents are yielded when those
        sheets are processed."""
        stack: List[MultiVariableBase] = [sheet_mv]
        while stack:
            mv = stack.pop()
            for name in mv._component_order:
                comp = mv._components[name]
                if isinstance(comp, Variable):
                    yield comp
                elif isinstance(comp, MultiVariableBase) and not comp._is_sheet:
                    stack.append(comp)

    def _collect_sheets(self) -> List[MultiVariableBase]:
        """Collect all MVs with _is_sheet=True from roots (depth-first)."""
        sheets: List[MultiVariableBase] = []
        for root in self.roots:
            if isinstance(root, MultiVariableBase):
                if root._is_sheet:
                    sheets.append(root)
                for _name, comp in root.walk():
                    if isinstance(comp, MultiVariableBase) and comp._is_sheet:
                        if comp not in sheets:
                            sheets.append(comp)
        return sheets
    
    def _layout_mv(self, mv: MultiVariableBase, row: int, col: int, sheet_name: str) -> int:
        """Recursively lay out an MV and its children.

        Returns the next available row after this MV's contents.
        """
        for name in mv._component_order:
            comp = mv._components[name]
            
            if isinstance(comp, MultiVariableBase):
                if comp._is_sheet:
                    continue
                section_id = comp.python_name or comp._name_in_parent or comp.id

                explicit_row = comp._creation_params.get("row")
                explicit_col = comp._creation_params.get("col")
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


