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
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .translator import ExcelTranslator
from .layout import LayoutEngine
from .addresses import VariableAddresses
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
    roots = _collect_root_sheets(root)
    if not roots:
        raise ValueError(
            "Nothing to emit: this MultiVariable has no content to put in the "
            "workbook. Add Variables or nested MultiVariables first, e.g.\n"
            "    pnl = mo.MultiVariable('P&L', excel_props={'tab': True})\n"
            "    pnl.revenue = mo.Variable(1_000_000)\n"
            "Then call to_excel() on the enclosing MultiVariable."
        )

    engine = LayoutEngine(roots)
    addresses = engine.compute_addresses()
    if not addresses:
        raise ValueError(
            "Tabs exist but contain no Variables. "
            "Every tab needs at least one Variable so the writer has "
            "something to write. Example:\n"
            "    inputs = mo.MultiVariable('Inputs', excel_props={'tab': True})\n"
            "    inputs.revenue = mo.Variable(1_000_000, display_name='Revenue')"
        )

    var_to_sheet = _build_var_to_sheet(addresses, roots)
    translator = ExcelTranslator(addresses, var_to_sheet)

    wb = Workbook()
    wb.remove(wb.active)

    for sheet_mv in roots:
        sheet_name = sheet_mv.sheet_name or sheet_mv.id
        ws = wb.create_sheet(title=_safe_sheet_name(sheet_name))
        _write_sheet(ws, sheet_mv, addresses, translator, sheet_name,
                     engine.section_header_rows)

    wb.save(path)


def _collect_root_sheets(root: "MultiVariableBase") -> list[MultiVariableBase]:
    """Collect sheets to emit from an explicit root.

    Resolution order:

    1. ``root`` has ``excel_props={'tab': True}`` → sole tab.
    2. ``root`` has sheet-roled descendants anywhere → those become the tabs.
    3. Fallback: no explicit sheet markers in the tree. Synthesize tabs
       from ``root``'s first-depth structure — each first-depth
       ``MultiVariable`` becomes a tab, and any direct Variables on
       ``root`` are wrapped in a synthetic "overview" tab so they still
       land in the workbook.
    """
    if root._is_sheet:
        return [root]

    # FIFO traversal so sheets land in declaration order (a stack with
    # ``pop()`` reverses every level).
    from collections import deque
    sheets: list[MultiVariableBase] = []
    queue = deque([root])
    while queue:
        mv = queue.popleft()
        if mv is not root and mv._is_sheet:
            sheets.append(mv)
            continue  # Sheets are leaves of the tab hierarchy — don't recurse.
        for child in mv._components.values():
            if isinstance(child, MultiVariableBase):
                queue.append(child)
    if sheets:
        return sheets

    return _synthesize_sheets(root)


def _synthesize_sheets(root: "MultiVariableBase") -> list[MultiVariableBase]:
    """Create virtual sheets when the user hasn't marked any ``MultiVariable``
    with ``excel_props={'tab': True}``. First-depth sub-MVs get promoted to sheets;
    direct Variables on ``root`` are collected into a single synthetic
    "overview" sheet named after ``root``.
    """
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
        out.append(_make_virtual_sheet(overview_name, direct_vars))

    for mv in sub_mvs:
        if mv._is_sheet:
            out.append(mv)
        else:
            # Promote a plain MV to a virtual Sheet without mutating the
            # user's model — new wrapper, same components by reference.
            label = mv.display_name or mv.python_name or mv.id
            items = [(n, mv._components[n]) for n in mv._component_order]
            out.append(_make_virtual_sheet(label, items))
    return out


def _make_virtual_sheet(
    name: str, components: list[tuple[str, Any]]
) -> MultiVariableBase:
    """Bare-bones Sheet-roled MV for layout / writing.

    Uses ``object.__new__`` so the synthesized sheet doesn't attach
    itself to an enclosing ``with`` context or mint a QPath — it's a
    render-time view, not a real model node.
    """
    from ...core.multi_variable import MultiVariable
    virt = object.__new__(MultiVariable)
    virt._role = 'sheet'
    virt._display_name = name
    virt._components = {}
    virt._component_order = []
    virt._parent = None
    virt._name_in_parent = None
    virt._qualified_id = None
    virt._python_name = None
    virt._creation_params = {}
    virt._excel_props = {'tab': True}
    for comp_name, comp in components:
        virt._components[comp_name] = comp
        virt._component_order.append(comp_name)
    return virt


def _build_var_to_sheet(
    addresses: dict[str, VariableAddresses], roots: list[MultiVariableBase]
) -> dict[str, str]:
    """Map each variable's qualified-path id to the sheet name it lives on."""
    var_to_sheet: dict[str, str] = {}
    for sheet_mv in roots:
        sheet_name = sheet_mv.sheet_name or sheet_mv.id
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


def _write_sheet(
    ws: Worksheet,
    sheet_mv: MultiVariableBase,
    addresses: dict[str, VariableAddresses],
    translator: ExcelTranslator,
    sheet_name: str,
    section_header_rows: dict[str, tuple],
) -> None:
    """Fill a worksheet: section-header labels for nested MVs, then a
    row per Variable (label + per-period values or formulas)."""
    # Section headers first — every nested (non-sheet) MV gets its
    # display_name written at the reserved (row, col) slot so the
    # Cohorts / Assumptions / etc. layout reads naturally.
    for _, comp in sheet_mv.walk():
        if isinstance(comp, MultiVariableBase) and not comp._is_sheet:
            section_id = comp.python_name or comp._name_in_parent or comp.id
            pos = section_header_rows.get(section_id)
            if pos is not None and comp.display_name:
                row, col = pos
                ws.cell(row=row, column=col, value=comp.display_name)

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
        number_format = comp._excel_props.get("number_format") if comp._excel_props else None

        for i, cell_addr in enumerate(addr.values):
            cell = ws[cell_addr]
            cell.value = _cell_content(comp, i, var_id, translator, sheet_name, addr)
            if number_format:
                cell.number_format = number_format

    _auto_width_first_column(ws)


def _cell_content(
    var: Variable,
    period_idx: int,
    var_id: str,
    translator: ExcelTranslator,
    current_sheet: str,
    self_address: VariableAddresses,
) -> Any:
    """Return the value to put in a cell: an Excel formula for computed variables, else the raw value."""
    if var._expr is not None:
        # Variables with an expression tree → translate to Excel formula.
        # Variables without one (plain inputs, lambda recurrence results) fall
        # through to value-emission below.
        try:
            excel = translator.translate(
                var._expr,
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
