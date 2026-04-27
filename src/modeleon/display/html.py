# SPDX-License-Identifier: Apache-2.0
"""HTML rendering for Jupyter ``_repr_html_`` methods.

Cells display Excel coordinates: column letters across the top
(``A``, ``B``, ``C``, …) and row numbers down the left side
(``1``, ``2``, ``3``, …), matching the addresses the workbook will
receive. Computed cells show their Excel formula on focus or via the
"Show formulas" toggle.

Notebook environments call ``obj._repr_html_()`` when displaying an
object. Modeleon objects render as compact tables — Variables become a
row, MultiVariables become a grid, Sheets and Models become titled
sections.

Each data cell carries **both** its computed value and the Excel
formula it compiles to. A CSS-only toggle button above the table
switches between the two views. Default view is values (more
scannable); click "Show formulas" to see the ``=B4*B5`` Excel output.

Styling uses the Modeleon brand palette (teal primary, gold accent,
navy text, green-gray borders — see assets/COLOR_PALETTE.md). All
styles are inline or in small scoped ``<style>`` blocks — no external
stylesheet — so rendering works uniformly across classic Jupyter,
JupyterLab, nbviewer, VSCode notebooks. The toggle is pure CSS (no
JavaScript), using an invisible checkbox + sibling selectors. A small
inline script adds trace-precedents highlighting on focus (Excel-like
"show me what this formula depends on") and degrades gracefully when
JS is sandboxed — the pure-CSS focus outline still reveals the
formula.
"""

from __future__ import annotations

import html
import itertools
import logging
import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Optional


logger = logging.getLogger(__name__)


# ─── Inline style fragments (quoted with double quotes in HTML attrs) ─
#
# Brand palette reference (see assets/COLOR_PALETTE.md):
#   Teal (primary)       #07464a
#   Teal (bright accent) #0d9a9f
#   Gold (alert/trace)   #bfa163
#   Navy (text)          #262c38
#   Text secondary       #4a5060
#   Text dim             #7a8080
#   Border               #c4cec6
#   Border subtle        #e2e8e4
#   Surface              #f7f8f7
#   Surface elevated     #f0f2f1

_TABLE = (
    "border-collapse:collapse; "
    "font-family:'Inter',-apple-system,BlinkMacSystemFont,system-ui,sans-serif; "
    "font-variant-numeric:tabular-nums; "
    "font-size:13px; "
    "color:#262c38; "
    # Explicit white background so dark-theme notebooks don't bleed
    # through the transparent TD cells.
    "background:#ffffff; "
    "border:1px solid #c4cec6; "
    "border-radius:4px; "
    "overflow:hidden; "
    "box-shadow:0 1px 2px rgba(38,44,56,0.04); "
    "margin:6px 0;"
)
_TH_LABEL = (
    "padding:6px 12px; background:#f7f8f7; border-bottom:1px solid #e2e8e4; "
    "border-right:1px solid #e2e8e4; text-align:left; font-weight:600; "
    "color:#262c38;"
)
_TH_HEADER = (
    "padding:6px 12px; background:#f7f8f7; border-bottom:1px solid #e2e8e4; "
    "border-right:1px solid #e2e8e4; text-align:right; font-weight:500; "
    "color:#4a5060;"
)
_TD_NUM = (
    "padding:6px 12px; background:#ffffff; "
    "border-right:1px solid #e2e8e4; "
    "border-bottom:1px solid #e2e8e4; text-align:right; color:#262c38;"
)
_TD_TXT = (
    "padding:6px 12px; background:#ffffff; "
    "border-right:1px solid #e2e8e4; "
    "border-bottom:1px solid #e2e8e4; text-align:left; color:#262c38;"
)
# Inline style applied to the .mo-formula span — monospace + brand teal.
# Single quotes are safe inside a double-quoted style attribute.
_FORMULA_INLINE = (
    "font-family:'JetBrains Mono','SF Mono',Menlo,monospace; "
    "font-size:11px; color:#07464a;"
)
_SECTION = (
    "padding:6px 12px; background:#f0f2f1; border-bottom:1px solid #e2e8e4; "
    "text-align:left; font-weight:600; color:#07464a; "
    "font-size:11px; letter-spacing:0.04em; text-transform:uppercase;"
)
# Excel column-letter header (top of each table) — subtle, monospace,
# centered. The axis label, not a focal point.
_TH_COL = (
    "padding:3px 12px; background:#f7f8f7; border-bottom:1px solid #e2e8e4; "
    "border-right:1px solid #e2e8e4; text-align:center; "
    "font-family:'JetBrains Mono','SF Mono',Menlo,monospace; "
    "font-size:10px; color:#7a8080; font-weight:500;"
)
# Excel row-number column (left of each table) — same subtle treatment.
_TH_ROW = (
    "padding:6px 10px; background:#f7f8f7; border-right:1px solid #e2e8e4; "
    "border-bottom:1px solid #e2e8e4; text-align:right; "
    "font-family:'JetBrains Mono','SF Mono',Menlo,monospace; "
    "font-size:10px; color:#7a8080; font-weight:500;"
)


# Running ID counter so multiple renders in one notebook don't collide
# on CSS selector IDs.
_uid_seq = itertools.count(1)


def _next_uid() -> str:
    return f"mo-repr-{next(_uid_seq)}"


# ─── Value formatting ─────────────────────────────────────────────


def _format_value(value: Any) -> str:
    """Pretty-format a value for HTML display."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int):
        return f"{value:,}"
    if isinstance(value, float):
        if value != value:  # NaN
            return "NaN"
        if abs(value) >= 1000:
            return f"{value:,.0f}"
        if abs(value) < 0.01 and value != 0:
            return f"{value:.4f}"
        return f"{value:,.2f}"
    if isinstance(value, (date, datetime)):
        return value.isoformat() if isinstance(value, date) else value.date().isoformat()
    return str(value)


def _is_text_value(value: Any) -> bool:
    """True for strings/dates/bools (left-aligned in cells); False for numbers."""
    if isinstance(value, (str, date, datetime)) or isinstance(value, bool):
        return True
    return False


def _label_with_unit(var) -> str:
    """Variable's display name with its unit suffix when present."""
    label = var.display_name
    unit = getattr(var, '_unit', None)
    if unit is not None:
        unit_str = str(unit)
        if unit_str:
            label = f"{label} ({unit_str})"
    return label


# ─── Translator context (formula-mode rendering) ──────────────────


@dataclass
class _Ctx:
    """Per-render context: address table + translator + var→sheet map.

    Built once per top-level ``_repr_html_`` call (model / sheet / mv)
    and threaded down to per-cell renders so we can show real Excel
    formulas without re-building the layout for every Variable.

    Failures during construction (e.g. an MV with no Sheet inside it)
    yield ``None`` — callers fall back to plain-value rendering.
    """
    addresses: dict
    translator: Any
    var_to_sheet: dict


def _build_ctx(root) -> Optional[_Ctx]:
    """Build a translator context for *root*. Returns ``None`` if the
    root has no Sheets to lay out (e.g. a bare MultiVariable used for
    grouping)."""
    try:
        from ..compile.excel.layout import LayoutEngine
        from ..compile.excel.translator import ExcelTranslator
        from ..compile.excel.writer import _build_var_to_sheet, _collect_root_sheets

        roots = _collect_root_sheets(root)
        if not roots:
            return None
        addresses = LayoutEngine(roots).compute_addresses()
        if not addresses:
            return None
        var_to_sheet = _build_var_to_sheet(addresses, roots)
        translator = ExcelTranslator(addresses, var_to_sheet)
        return _Ctx(addresses=addresses, translator=translator, var_to_sheet=var_to_sheet)
    except Exception as exc:
        logger.debug("HTML repr: translator context unavailable (%s)", exc)
        return None


def _parse_addr(addr: str) -> tuple[str, int]:
    """Split an Excel address like ``B5`` into ``("B", 5)``.

    Handles single- and multi-letter columns (``AA42`` → ``("AA", 42)``).
    Returns ``("", 0)`` for malformed input.
    """
    i = 0
    while i < len(addr) and addr[i].isalpha():
        i += 1
    if i == 0 or i == len(addr):
        return "", 0
    try:
        return addr[:i], int(addr[i:])
    except ValueError:
        return "", 0


def _value_col_letters(var, ctx: Optional[_Ctx]) -> list[str]:
    """Excel column letters for each value cell of *var*. Empty list
    if no layout context or no addresses recorded for this Variable."""
    if ctx is None:
        return []
    addr_obj = ctx.addresses.get(var.id)
    if addr_obj is None:
        return []
    return [_parse_addr(a)[0] for a in addr_obj.values]


def _var_row(var, ctx: Optional[_Ctx]) -> Optional[int]:
    """Excel row number for *var*. ``None`` if no layout context."""
    if ctx is None:
        return None
    addr_obj = ctx.addresses.get(var.id)
    if addr_obj is None:
        return None
    return _parse_addr(addr_obj.name)[1] or None


# Cell-reference regex used to extract precedents from a rendered
# Excel formula. Captures three shapes:
#   - 'Sheet Name'!B5  (quoted sheet, single cell or A:B range)
#   - SheetName!B5     (unquoted sheet)
#   - B5               (local cell or range, no sheet prefix)
# Range form is captured via the optional ``:colN rowN`` tail.
_CELL_REF = re.compile(
    r"(?:'(?P<qsheet>[^']+)'|(?P<sheet>[A-Za-z_]\w*))!"
    r"(?P<col1>[A-Z]+)(?P<row1>\d+)"
    r"(?::(?P<col2>[A-Z]+)(?P<row2>\d+))?"
    r"|(?P<lcol1>[A-Z]+)(?P<lrow1>\d+)"
    r"(?::(?P<lcol2>[A-Z]+)(?P<lrow2>\d+))?"
)


def _col_letter_to_int(col: str) -> int:
    """Excel column letters to 1-based int. ``A`` → 1, ``Z`` → 26,
    ``AA`` → 27, ``AZ`` → 52, ``BA`` → 53."""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n


def _int_to_col_letter(n: int) -> str:
    """Inverse of :func:`_col_letter_to_int`."""
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(r + ord('A')) + out
    return out


def _expand_range(col1: str, row1: int, col2: str, row2: int) -> list[tuple[str, int]]:
    """Expand an A1-style range like ``B5:D7`` into a list of (col_letter, row)
    pairs. Order: rows first within each column."""
    c1, c2 = _col_letter_to_int(col1), _col_letter_to_int(col2)
    if c1 > c2:
        c1, c2 = c2, c1
    if row1 > row2:
        row1, row2 = row2, row1
    cells = []
    for c in range(c1, c2 + 1):
        col = _int_to_col_letter(c)
        for r in range(row1, row2 + 1):
            cells.append((col, r))
    return cells


def _extract_precedent_cells(formula: str, current_sheet: str) -> list[str]:
    """Parse an Excel formula and return the set of fully-qualified
    (``Sheet!Address``) cell addresses it references. Ranges are
    expanded to individual cells. Cell refs without a sheet prefix
    are qualified with ``current_sheet``.

    De-duplicated; preserves first-encounter order so JS highlights
    happen in formula-reading order (mostly cosmetic).
    """
    seen: set[str] = set()
    out: list[str] = []
    for m in _CELL_REF.finditer(formula or ""):
        if m.group("col1") is not None:
            sheet = m.group("qsheet") or m.group("sheet") or current_sheet
            col1, row1 = m.group("col1"), int(m.group("row1"))
            col2 = m.group("col2") or col1
            row2 = int(m.group("row2") or row1)
        elif m.group("lcol1") is not None:
            sheet = current_sheet
            col1, row1 = m.group("lcol1"), int(m.group("lrow1"))
            col2 = m.group("lcol2") or col1
            row2 = int(m.group("lrow2") or row1)
        else:
            continue
        if not sheet:
            continue
        for c, r in _expand_range(col1, row1, col2, row2):
            addr = f"{sheet}!{c}{r}"
            if addr not in seen:
                seen.add(addr)
                out.append(addr)
    return out


def _excel_formula_for(var, period_idx: int, ctx: _Ctx) -> Optional[str]:
    """Render a Variable's Excel formula at ``period_idx``, or None if
    the Variable has no expression / no layout address / translation
    fails. The leading ``=`` is included for cells that compile to
    formulas."""
    if var._expr is None:
        return None
    addr = ctx.addresses.get(var.id)
    if addr is None or period_idx >= len(addr.values):
        return None
    sheet_name = ctx.var_to_sheet.get(var.id)
    try:
        excel = ctx.translator.translate(
            var._expr,
            period_idx,
            current_sheet=sheet_name,
            self_address=addr,
            self_var=var,
        )
    except Exception as exc:
        logger.debug("HTML repr: translate failed for %s: %s", var.path, exc)
        return None
    if not excel:
        return None
    excel_str = str(excel)
    return excel_str if excel_str.startswith("=") else f"={excel_str}"


# ─── Toggle + tab wrappers (CSS-only) ──────────────────────────────


# Root container — light surface wrapping everything so the repr
# stays readable on dark-theme notebooks. Shared across toggle- and
# tabbed-wrappers. ``display:inline-block`` shrink-wraps to fit content
# (avoids stretching to the full notebook width on narrow tables),
# while ``max-width:100%`` caps at the parent so wide tables scroll
# internally instead of overflowing the container's rounded border.
_CONTAINER_STYLE = (
    "display:inline-block; "
    "background:#ffffff; color:#262c38; "
    "padding:12px 14px; border:1px solid #c4cec6; border-radius:6px; "
    "box-shadow:0 1px 3px rgba(38,44,56,0.05); "
    "max-width:100%; box-sizing:border-box; overflow:hidden;"
)

# Ghost-style toolbar button — no border, subtle text, hover shows a
# light surface tint, active state fills with cream + teal text (gold-
# adjacent, not shouting). Works for both the formula toggle and the
# stack-sheets toggle via the `.mo-btn` class.
_TOGGLE_STYLES = (
    "{root} .mo-formula {{ display: none; }}"
    "{root} input.mo-toggle {{ display: none; }}"
    "{root} input.mo-toggle:checked ~ * .mo-val {{ display: none; }}"
    "{root} input.mo-toggle:checked ~ * .mo-formula {{ display: inline; }}"
    # Toolbar sits to the right of the tabs (or on its own above a
    # single-sheet table).
    "{root} .mo-toolbar {{"
    "display: inline-flex; gap: 2px; align-items: center; "
    "margin-left: auto;"
    "}}"
    "{root} label.mo-btn {{"
    "cursor: pointer; padding: 4px 10px; background: transparent; "
    "border: none; border-radius: 4px; font-size: 11px; "
    "font-family: 'Inter',system-ui,sans-serif; color: #7a8080; "
    "font-weight: 500; display: inline-flex; align-items: center; "
    "user-select: none; letter-spacing: 0.02em; "
    "transition: background 120ms ease, color 120ms ease;"
    "}}"
    "{root} label.mo-btn:hover {{ background: #f0f2f1; color: #07464a; }}"
    "{root} input.mo-toggle:checked ~ * label.mo-btn.mo-formula-btn, "
    "{root} input.mo-toggle:checked ~ label.mo-btn.mo-formula-btn "
    "{{ background: #fffaf2; color: #07464a; }}"
    "{root} label.mo-btn.mo-formula-btn::before "
    "{{ content: '\u0192\u00a0'; font-weight: 600; color: #bfa163; margin-right: 2px; }}"
    "{root} label.mo-btn.mo-formula-btn::after {{ content: 'Formulas'; }}"
    "{root} input.mo-toggle:checked ~ * label.mo-btn.mo-formula-btn::after, "
    "{root} input.mo-toggle:checked ~ label.mo-btn.mo-formula-btn::after "
    "{{ content: 'Values'; }}"
    # Per-cell click-to-inspect: focusing a formula cell shows its
    # formula. Outline uses bright teal so the active cell pops.
    "{root} td.mo-cell {{ cursor: pointer; transition: background 80ms ease; }}"
    "{root} td.mo-cell:hover {{ background: #f7f8f7; }}"
    "{root} td.mo-cell:focus {{"
    "outline: 2px solid #0d9a9f; outline-offset: -2px; background: #ffffff;"
    "}}"
    "{root} td.mo-cell:focus .mo-val {{ display: none; }}"
    "{root} td.mo-cell:focus .mo-formula {{ display: inline; }}"
    # Trace-precedents — gold dashed outline + soft-cream tint.
    # !important needed because cells carry inline ``style="background:#ffffff"``
    # (numeric/text td defaults), and inline styles outrank class selectors.
    "{root} td.mo-dep {{"
    "outline: 1px dashed #bfa163 !important; outline-offset: -1px; "
    "background: #fffaf2 !important;"
    "}}"
)


_TAB_STYLES = (
    "{root} input.mo-tab {{ display: none; }}"
    "{root} .mo-panel {{ display: none; }}"
    # Toolbar lives in its own row above the tab bar, left-aligned.
    # Transparent background + different vertical position means these
    # buttons read as chrome/controls rather than more tabs.
    "{root} .mo-toolbar {{"
    "display: flex; justify-content: flex-start; align-items: center; "
    "gap: 2px; margin: 2px 0 4px 0; min-height: 24px;"
    "}}"
    # Tab bar is a plain strip of labels — the Excel-like row below
    # the toolbar. Light surface + border ties it to the panel below.
    "{root} .mo-tab-bar {{"
    "display: flex; align-items: stretch; gap: 2px; "
    "background: #f7f8f7; border: 1px solid #c4cec6; "
    "border-bottom: none; border-radius: 4px 4px 0 0; "
    "padding: 2px 4px 0 4px; overflow-x: auto;"
    "}}"
    "{root} .mo-tab-bar label.mo-tab-label {{"
    "cursor: pointer; padding: 6px 14px; font-size: 12px; "
    "font-family: 'Inter',system-ui,sans-serif; color: #4a5060; "
    "background: transparent; border-radius: 4px 4px 0 0; "
    "user-select: none; position: relative; bottom: -1px; "
    "white-space: nowrap; "
    "transition: color 120ms ease, background 120ms ease;"
    "}}"
    "{root} .mo-tab-bar label.mo-tab-label:hover "
    "{{ background: #ffffff; color: #07464a; }}"
    # The sheets panel below sits flush with the tab bar — no top gap,
    # matching border color. ``overflow-x: auto`` lets wide tables
    # (12-period time series, etc.) scroll horizontally inside the
    # panel instead of overflowing the container's rounded border.
    "{root} .mo-panels {{"
    "border: 1px solid #c4cec6; border-top: none; "
    "border-radius: 0 0 4px 4px; background: #ffffff; padding: 8px; "
    "overflow-x: auto;"
    "}}"
    # Per-panel title: hidden in tabbed mode (tab label is the header),
    # revealed in stacked mode so each section is identifiable.
    "{root} .mo-panel-title {{ display: none; }}"
    # Space stacked panels vertically so they don't run together.
    "{root} .mo-panel + .mo-panel {{ margin-top: 14px; }}"
    # Stacked-view toggle: show all panels + hide just the tab bar when
    # checked. Toolbar stays — it lives above in its own row.
    "{root} input.mo-stack {{ display: none; }}"
    "{root} input.mo-stack:checked ~ .mo-panels .mo-panel "
    "{{ display: block; }}"
    "{root} input.mo-stack:checked ~ .mo-panels .mo-panel-title "
    "{{ display: block; }}"
    "{root} input.mo-stack:checked ~ .mo-tab-bar "
    "{{ display: none; }}"
    # Without the tab bar, the panels would lose their top rounded
    # corners and top border — restore them.
    "{root} input.mo-stack:checked ~ .mo-panels "
    "{{ border-top: 1px solid #c4cec6; border-radius: 4px; }}"
    "{root} label.mo-btn.mo-stack-btn::before "
    "{{ content: '\u25a6\u00a0'; font-weight: 600; color: #bfa163; margin-right: 2px; }}"
    "{root} label.mo-btn.mo-stack-btn::after {{ content: 'Stacked'; }}"
    "{root} input.mo-stack:checked ~ * label.mo-btn.mo-stack-btn, "
    "{root} input.mo-stack:checked ~ label.mo-btn.mo-stack-btn "
    "{{ background: #fffaf2; color: #07464a; }}"
    "{root} input.mo-stack:checked ~ * label.mo-btn.mo-stack-btn::after, "
    "{root} input.mo-stack:checked ~ label.mo-btn.mo-stack-btn::after "
    "{{ content: 'Tabbed'; }}"
)


def _precedents_script(uid: str) -> str:
    """Inline script for Excel-like trace-precedents on focus.

    When the user clicks or Tabs into a formula cell, every cell the
    formula references gets a dashed gold outline + cream background —
    the same visual idea as Excel's "Trace Precedents" button. Losing
    focus clears the highlight.

    Pure-CSS focus-outline and click-to-show-formula already work
    without this script; it's a progressive enhancement. If JS is
    sandboxed (some JupyterLab configs), the toggle button and
    per-cell click-to-inspect still function.
    """
    # Use a pre-built address→cell map instead of querySelector. CSS
    # attribute selectors with characters like ``&`` (sheet names: "P&L")
    # are spec-legal but not universally supported by all engines, and
    # jsdom in particular silently fails them. The map sidesteps all
    # encoding edge cases and is faster (O(deps) per click).
    return (
        f"<script>(function(){{"
        f"var w=document.getElementById('{uid}');if(!w)return;"
        f"var idx={{}};"
        f"w.querySelectorAll('[data-cell]').forEach("
        f"function(e){{idx[e.getAttribute('data-cell')]=e;}});"
        f"function clr(){{w.querySelectorAll('td.mo-dep').forEach("
        f"function(e){{e.classList.remove('mo-dep');}});}}"
        f"w.addEventListener('focusin',function(e){{"
        f"clr();"
        f"var cell=e.target.closest&&e.target.closest('td.mo-cell');"
        f"if(!cell)return;"
        f"var deps;try{{deps=JSON.parse(cell.dataset.deps||'[]');}}"
        f"catch(_){{deps=[];}}"
        f"deps.forEach(function(addr){{"
        f"var dep=idx[addr];"
        f"if(dep)dep.classList.add('mo-dep');}});}});"
        f"w.addEventListener('focusout',function(e){{"
        f"if(!w.contains(e.relatedTarget))clr();}});"
        f"}})();</script>"
    )


def _toggle_wrapper(uid: str, body: str, has_any_formula: bool) -> str:
    """Wrap *body* in a light-surface container with an optional formula-view
    toggle in a small top-right toolbar.

    If no Variable in the body has a formula, the toolbar is omitted —
    there's nothing to toggle. Either way, the container keeps the
    repr readable on dark-theme notebooks.
    """
    if not has_any_formula:
        return (
            f'<div id="{uid}" style="{_CONTAINER_STYLE}">{body}</div>'
        )
    root = f"#{uid}"
    toolbar = (
        f'<div class="mo-toolbar" style="display:flex;justify-content:flex-start;margin-bottom:4px;">'
        f'<label for="{uid}-cb" class="mo-btn mo-formula-btn"></label>'
        f'</div>'
    )
    return (
        f"<style>{_TOGGLE_STYLES.format(root=root)}</style>"
        f'<div id="{uid}" tabindex="0" style="{_CONTAINER_STYLE}">'
        f'<input type="checkbox" class="mo-toggle" id="{uid}-cb">'
        f"{toolbar}"
        f"<div>{body}</div>"
        f"</div>"
        f"{_precedents_script(uid)}"
    )


def _tabbed_wrapper(
    uid: str,
    panels: list[tuple[str, str]],
    has_any_formula: bool,
    title: str = "",
) -> str:
    """Excel-style tab UI over a list of (tab_label, panel_html) pairs.

    Each panel is a separate sheet's table. Hidden radio buttons (one
    per sheet, all in the same named group) drive which panel is
    visible via `input:checked ~ .mo-panels .mo-panel-N` selectors.
    Formula-toggle + stack-sheets buttons live in a toolbar on the
    right side of the tab bar — toolbar feel rather than floating
    pills above. Everything wraps in a light-surface container so the
    repr reads cleanly on dark-theme notebooks.
    """
    root = f"#{uid}"
    n = len(panels)

    # Per-sheet CSS: active tab lifts to white + teal text, matching
    # panel becomes visible. Reads against the gray tab-bar background.
    tab_css_parts = []
    for i in range(n):
        tab_id = f"{uid}-t{i}"
        tab_css_parts.append(
            f"{root} #{tab_id}:checked ~ .mo-tab-bar label[for='{tab_id}']"
            f"{{ background: #ffffff; color: #07464a; font-weight: 600; "
            f"box-shadow: 0 -1px 0 #c4cec6 inset, -1px 0 0 #c4cec6 inset, "
            f"1px 0 0 #c4cec6 inset; }}"
        )
        tab_css_parts.append(
            f"{root} #{tab_id}:checked ~ .mo-panels .mo-panel-{i}"
            f"{{ display: block; }}"
        )
    per_sheet_css = "".join(tab_css_parts)

    style = (
        f"<style>"
        f"{_TOGGLE_STYLES.format(root=root) if has_any_formula else ''}"
        f"{_TAB_STYLES.format(root=root)}"
        f"{per_sheet_css}"
        f"</style>"
    )

    radios = "".join(
        f'<input type="radio" class="mo-tab" name="{uid}-tabs" '
        f'id="{uid}-t{i}"{" checked" if i == 0 else ""}>'
        for i in range(n)
    )
    stacked_toggle = (
        f'<input type="checkbox" class="mo-stack" id="{uid}-stack">'
    )
    formula_toggle = (
        f'<input type="checkbox" class="mo-toggle" id="{uid}-cb">'
        if has_any_formula else ""
    )

    tab_labels = "".join(
        f'<label for="{uid}-t{i}" class="mo-tab-label">{html.escape(label)}</label>'
        for i, (label, _) in enumerate(panels)
    )
    panel_divs = "".join(
        f'<div class="mo-panel mo-panel-{i}">{panel_html}</div>'
        for i, (_, panel_html) in enumerate(panels)
    )

    toolbar_buttons = ""
    if has_any_formula:
        toolbar_buttons += (
            f'<label for="{uid}-cb" class="mo-btn mo-formula-btn"></label>'
        )
    toolbar_buttons += (
        f'<label for="{uid}-stack" class="mo-btn mo-stack-btn"></label>'
    )
    toolbar = f'<div class="mo-toolbar">{toolbar_buttons}</div>'

    precedents_script = _precedents_script(uid) if has_any_formula else ""
    return (
        f"{style}"
        f'<div id="{uid}" tabindex="0" style="{_CONTAINER_STYLE}">'
        f"{title}"
        f"{radios}"
        f"{stacked_toggle}"
        f"{formula_toggle}"
        f'{toolbar}'
        f'<div class="mo-tab-bar">{tab_labels}</div>'
        f'<div class="mo-panels">{panel_divs}</div>'
        f"</div>"
        f"{precedents_script}"
    )


# ─── Variable rendering ───────────────────────────────────────────


def variable_html(var) -> str:
    """Render a single Variable as an HTML table row."""
    ctx = _build_ctx(var._owner) if getattr(var, '_owner', None) is not None else None
    label = _label_with_unit(var)
    value = var._value
    values = value if isinstance(value, list) else [value]
    cells_and_flags = [_render_var_cell(var, v, i, ctx) for i, v in enumerate(values)]
    has_formula = any(has for _, has in cells_and_flags)
    cells = "".join(cell for cell, _ in cells_and_flags)
    table = (
        f'<table style="{_TABLE}">'
        f'<tr><th style="{_TH_LABEL}">{html.escape(str(label))}</th>{cells}</tr>'
        f"</table>"
    )
    return _toggle_wrapper(_next_uid(), table, has_any_formula=has_formula)


def _render_var_cell(
    var, value, period_idx: int, ctx: Optional[_Ctx]
) -> tuple[str, bool]:
    """Return (cell_html, has_formula) for a single Variable cell.

    For computed cells with a formula, the cell contains two spans:
    ``.mo-val`` (the value) and ``.mo-formula`` (the Excel formula).
    The toggle CSS swaps which is visible. Input cells contain only
    the value (nothing to toggle).

    Every cell carries ``data-cell="Sheet!Addr"`` (its sheet-qualified
    Excel address). Formula cells also carry ``data-deps`` listing the
    sheet-qualified addresses of the **specific cells this formula
    references** (parsed from the rendered formula string). The focus
    handler highlights only those precise cells — Excel's
    trace-precedents UX, not "highlight every cell of every dep
    Variable."
    """
    formatted_value = _format_value(value)
    formula = (
        _excel_formula_for(var, period_idx, ctx) if ctx is not None else None
    )
    td_style = _TD_TXT if _is_text_value(value) else _TD_NUM

    # Self-address attribute: every cell gets one so dep-highlight can
    # find it. Falls back to nothing if no layout context.
    cell_attrs = ""
    current_sheet = None
    if ctx is not None:
        addr_obj = ctx.addresses.get(var.id)
        current_sheet = ctx.var_to_sheet.get(var.id)
        if addr_obj is not None and current_sheet is not None and period_idx < len(addr_obj.values):
            self_addr = f"{current_sheet}!{addr_obj.values[period_idx]}"
            cell_attrs = f' data-cell="{html.escape(self_addr)}"'

    if formula is not None:
        deps = (
            _extract_precedent_cells(formula, current_sheet)
            if current_sheet else []
        )
        # Sheet names can contain spaces, parens, even ``|`` — no single
        # delimiter is safe. Encode as a JSON array so the JS side just
        # ``JSON.parse``s it.
        import json as _json
        deps_attr = (
            f" data-deps='{html.escape(_json.dumps(deps))}'" if deps else ""
        )
        # ``tabindex="0"`` makes the cell keyboard-focusable. Combined
        # with the ``td.mo-cell:focus`` CSS rule above, click or Tab
        # flips it from value → formula display (Excel F2 equivalent).
        cell = (
            f'<td class="mo-cell" tabindex="0" style="{td_style}"'
            f"{cell_attrs}{deps_attr}>"
            f'<span class="mo-val">{html.escape(formatted_value)}</span>'
            f'<span class="mo-formula" style="{_FORMULA_INLINE}">'
            f"{html.escape(formula)}</span>"
            f"</td>"
        )
        return cell, True
    return (
        f'<td style="{td_style}"{cell_attrs}>'
        f"{html.escape(formatted_value)}</td>"
    ), False


# ─── MultiVariable rendering ──────────────────────────────────────


def _collect_rows(mv, direct_only: bool = False) -> tuple[list, int]:
    """Walk an MV's components, yielding (kind, label, var_or_None, values) rows.

    Each ``var`` row carries the Variable instance itself so per-cell
    formula rendering can look it up in the translator context.

    Returns the row list plus the maximum column count across all
    list-valued Variables (1 if everything is scalar).

    ``direct_only=True`` skips sub-MVs entirely (used by the synthetic
    first tab in :func:`model_html` — sub-MVs get their own tabs, so
    rendering them as nested sections here would double-up).
    """
    from ..core.variable import Variable
    from ..core.multi_variable import MultiVariableBase

    rows: list = []
    max_cols = 1
    for name in mv._component_order:
        comp = mv._components[name]
        if isinstance(comp, Variable):
            label = _label_with_unit(comp)
            values = comp._value if isinstance(comp._value, list) else [comp._value]
            max_cols = max(max_cols, len(values))
            rows.append(('var', label, comp, values))
        elif isinstance(comp, MultiVariableBase) and not comp._is_sheet:
            if direct_only:
                continue
            rows.append(('section', comp.display_name, None, None))
            sub_rows, sub_cols = _collect_rows(comp)
            max_cols = max(max_cols, sub_cols)
            rows.extend(sub_rows)
    return rows, max_cols


def _render_mv_table(
    mv,
    ctx: Optional[_Ctx],
    direct_only: bool = False,
    suppress_title: bool = False,
) -> tuple[str, bool]:
    """Build the inner table HTML for a MultiVariable + whether it
    contains at least one computed cell (to decide whether the
    toggle wrapper should show the button).

    When ``suppress_title`` is True, skip the leading ``display_name``
    caption — the caller is showing the MV inside a tabbed layout
    where the tab label already identifies it, so rendering the
    title here would duplicate it.

    Layout (with Excel axis):

        ┌──────┬───────┬─────┬─────┬─────┐
        │      │   A   │  C  │  D  │  E  │   ← column letters
        │      │       │  1  │  2  │  3  │   ← period header (if list)
        ├──────┼───────┼─────┼─────┼─────┤
        │  5   │ Label │ 100 │ 200 │ 300 │   ← row number + Variable row
        │  6   │ Label │ ... │ ... │ ... │
        └──────┴───────┴─────┴─────┴─────┘
    """
    rows, max_cols = _collect_rows(mv, direct_only=direct_only)
    if not rows:
        empty = html.escape(mv.display_name or 'MultiVariable')
        return f"<i>{empty} (empty)</i>", False

    # Discover Excel column letters for the value columns by sampling
    # the *widest* list-valued variable — the layout engine starts
    # every Variable at ``start_col`` so sharing columns is safe, but
    # sampling a scalar variable would only yield the first letter and
    # we'd pad the rest with blanks (no visible C, D, ...).
    wide_var = None
    wide_len = 0
    for kind, _, var, values in rows:
        if kind == 'var' and values is not None and len(values) > wide_len:
            wide_var = var
            wide_len = len(values)
    value_cols = _value_col_letters(wide_var, ctx) if wide_var else []
    # Fill or trim to ``max_cols`` so empty trailing columns still get
    # *some* letter (placeholder).
    if len(value_cols) < max_cols:
        value_cols = value_cols + [""] * (max_cols - len(value_cols))
    else:
        value_cols = value_cols[:max_cols]

    # Header rows: Excel column letters (always), then period labels
    # (only when there's more than one column).
    col_letter_cells = "".join(
        f'<th style="{_TH_COL}">{html.escape(c)}</th>' for c in value_cols
    )
    col_letter_row = (
        f'<tr><th style="{_TH_COL}"></th>'  # corner
        f'<th style="{_TH_COL}">A</th>'     # label column = Excel A
        f"{col_letter_cells}</tr>"
    )

    period_row = ""

    body = ""
    has_formula = False
    for kind, label, var, values in rows:
        if kind == 'section':
            body += (
                f'<tr><th style="{_TH_ROW}"></th>'
                f'<th style="{_SECTION}" colspan="{max_cols + 1}">'
                f"{html.escape(str(label))}</th></tr>"
            )
            continue
        cells = ""
        for i in range(max_cols):
            if values is not None and i < len(values):
                cell, has = _render_var_cell(var, values[i], i, ctx)
                cells += cell
                has_formula = has_formula or has
            else:
                cells += f'<td style="{_TD_NUM}"></td>'
        row_num = _var_row(var, ctx)
        row_label = str(row_num) if row_num else ""
        body += (
            f'<tr>'
            f'<th style="{_TH_ROW}">{row_label}</th>'
            f'<th style="{_TH_LABEL}">{html.escape(str(label))}</th>'
            f"{cells}</tr>"
        )

    title = (
        f'<div class="mo-panel-title" '
        f'style="font-family:\'Inter\',system-ui,sans-serif;'
        f"font-weight:600;font-size:13px;color:#07464a;"
        f'margin:2px 0 6px 0;">'
        f"{html.escape(mv.display_name)}</div>"
        if mv.display_name and not suppress_title else ""
    )
    table = (
        f"{title}"
        f'<table style="{_TABLE}">'
        f"{col_letter_row}"
        f"{period_row}"
        f"{body}"
        f"</table>"
    )
    return table, has_formula


def multi_variable_html(mv, ctx: Optional[_Ctx] = None) -> str:
    """Render a MultiVariable (or Sheet) as a titled grid.

    When called as a top-level ``_repr_html_``, ``ctx`` is None and
    this function builds its own translator context AND wraps the
    output in a toggle button. When called from ``model_html`` (which
    builds ctx once for the whole model and emits its own toggle
    wrapper), ``ctx`` is passed in and no toggle is added here.
    """
    is_top_level = ctx is None
    if ctx is None:
        ctx = _build_ctx(mv)

    table, has_formula = _render_mv_table(mv, ctx)
    if is_top_level:
        return _toggle_wrapper(_next_uid(), table, has_any_formula=has_formula)
    return table


# ─── Model rendering ──────────────────────────────────────────────


def model_html(root) -> str:
    """Render any MultiVariable as an Excel-style tabbed view.

    - First-depth sub-MultiVariables → one tab each.
    - Direct Variables on ``root`` → collected into an unnamed first tab.
    - Single tab → plain table without the tab bar.
    - Multiple tabs → tabbed interface with shared formula toggle.

    The translator context is built once for the whole tree so cross-sheet
    formula references (``=Assumptions!B3``) resolve correctly; rendering
    each panel in isolation would only know its own addresses.
    """
    from ..core.variable import Variable
    from ..core.multi_variable import MultiVariableBase

    ctx = _build_ctx(root)
    title = (
        f'<div style="font-family:\'Inter\',system-ui,sans-serif;'
        f"font-weight:700;font-size:15px;color:#07464a;"
        f'margin:4px 0 8px 0;letter-spacing:-0.01em;">'
        f"{html.escape(root.display_name or '')}</div>"
        if root.display_name else ""
    )

    direct_var_names: list = []
    sub_mvs: list = []
    for name in root._component_order:
        comp = root._components[name]
        if isinstance(comp, Variable):
            direct_var_names.append(name)
        elif isinstance(comp, MultiVariableBase):
            sub_mvs.append((name, comp))

    if not direct_var_names and not sub_mvs:
        return f"{title}<i>(empty)</i>"

    # Track which panels came from real sub-MVs (auto-promoted tabs in
    # Excel) vs. the synthetic overview of direct Variables on the root.
    # A single sub-MV panel still wants the Excel-tab visual; a single
    # synthetic overview panel doesn't (it'd label the tab with the root's
    # own name, which reads as redundant).
    panels_and_flags: list[tuple[str, str, bool, bool]] = []

    if direct_var_names:
        table, has_f = _render_mv_table(root, ctx, direct_only=True)
        overview_label = (
            root.display_name
            or getattr(root, 'python_name', None)
            or "Sheet1"
        )
        panels_and_flags.append((overview_label, table, has_f, False))

    for name, mv in sub_mvs:
        table, has_f = _render_mv_table(mv, ctx)
        label = mv.display_name or getattr(mv, 'python_name', None) or name
        panels_and_flags.append((label, table, has_f, True))

    has_any_formula = any(has for _, _, has, _ in panels_and_flags)

    panels = [(label, table) for label, table, _, _ in panels_and_flags]
    return _tabbed_wrapper(
        _next_uid(), panels, has_any_formula=has_any_formula, title=title,
    )
