# SPDX-License-Identifier: Apache-2.0
"""The blend row, written as FORMULAS — the splice stays alive in Excel.

A blended line ("live") is not data: it is ``given`` up to the close
boundary and ``follow`` after it, per cell, with a fallback to the
other side wherever one is missing. Emitted as numbers it would be
frozen in the written workbook — enter an actual in the file and the
blended line would sit unchanged, which is exactly the dead-values
habit Excel parity exists to kill.

So the blend row references its own track subrows:

* before the boundary — ``=IF(actual="", plan, actual)``;
* after it            — ``=IF(plan="", actual, plan)``;
* one side only       — a bare ``=<that cell>``.

Enter an actual in the workbook and the blended line moves with it,
exactly as the engine would have re-blended.

``actual=""`` asks "was an actual entered?" of a row of entries, which
is blank exactly where nothing was entered. A line can enter its plan
and DERIVE its actual (``mo.Variable(a + b, plan=[...])`` - the actual
of the sum is the sum of the actuals), and then its actual row is a
formula, a number even over months nobody entered (a blank counts as
zero). The engine's blend asks such a side whether any of the entries it
is built from is filled (``core.entries``); the book asks the same cells:
``=IF(COUNTA(<a's actual>,<b's actual>)=0,forecast,actual)``.

Only DATA lines take these — the rows ``expand_tracked_tree`` marked
``_blend_is_splice``, using the engine's own predicate (a line with
authored role kwargs, or one with no expression at all). A purely
DERIVED line's blend row already carries its own formula over the
operands' blend rows (a derived line's live series is computed from
its operands' live series, never spliced), and rewriting it would
replay the splice on an output.

A line can be both: ``mo.Variable(<plan formula>, actual=[...],
forecast=[...])`` computes its plan track from an expression and takes
the other two as authored series. It is a DATA line — its blend is a
splice of the authored rows — so the base expression is the wrong
formula for the blend row. Left there, the blend row would show plan
where the engine computes the splice, and an actual entered in the file
would never reach it.
"""

from typing import Any, Dict, List, Optional, Tuple

from .addresses import VariableAddresses
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable

#: The subrow attribute infix ``expand_tracked_tree`` emits.
TRACK_INFIX = "__track_"

#: Excel takes at most 255 arguments in one call, and 8192 characters
#: in one formula - the count leaves room for the rest of the splice.
_MAX_ARGS = 255
_MAX_COUNT_CHARS = 7000


def _before_boundary(sheet_mv: MultiVariableBase, spec: Any,
                     n: int) -> List[bool]:
    """Which periods take ``given`` first — the same splice the engine
    computes, spelled over the sheet's own window."""
    from ...core.time import _parse, _period_labels, resolve_default_window

    if spec.until is None:
        # Boundary-less: given wins wherever it carries a value.
        return [True] * n
    window = resolve_default_window(sheet_mv)
    if window is None or window.start is None or window.grain is None:
        return [True] * n
    labels = _period_labels(window.start, None, n, window.grain)
    try:
        anchor = _parse(spec.until, window.grain)
    except (ValueError, TypeError):
        return [True] * n
    return [_parse(lb, window.grain) <= anchor for lb in labels]


def _side_entries(head: Variable, side: Optional[Variable], role: str,
                  addresses: Dict[str, VariableAddresses],
                  tracked: set) -> Optional[List[str]]:
    """The rows of entries a DERIVED side of a splice is computed from
    (``core.entries`` - the walk the engine's blend asks), or None when
    the side is itself a row of entries (``=""`` asks it directly), reads
    none, or reads a row the book does not lay out."""
    if side is None:
        return None
    expr = getattr(side, "_expr", None)
    own = expr is not None                # the track's own expression
    if expr is None:
        if getattr(side, "_authored_track", None) == role:
            return None               # entered
        if getattr(head, "_expr_is_track_specific", False):
            return None
        expr = getattr(head, "_expr", None)
        if expr is None or any(
                dep in tracked and f"{dep}{TRACK_INFIX}{role}" not in addresses
                for dep in head.dependencies):
            return None               # the track writer keeps its numbers
    from ...core.entries import entry_lines
    try:
        entries = entry_lines(expr, role, own=own)
    except RecursionError:            # pragma: no cover - a book never dies here
        return None
    rows = [f"{line.id}{TRACK_INFIX}{track}" if track is not None else line.id
            for line, track in entries]
    if not rows or any(row not in addresses for row in rows):
        return None
    return rows


def _qualified(ref: str, sheet: Optional[str], here: str) -> str:
    if not sheet or sheet == here:
        return ref
    from .renderer import _UNQUOTED_SHEET_NAME
    if _UNQUOTED_SHEET_NAME.match(sheet):
        return f"{sheet}!{ref}"
    return "'" + sheet.replace("'", "''") + f"'!{ref}"


def _entered(rows: List[str], i: int, n: int, addresses: Dict[str, VariableAddresses],
             var_to_sheet: Dict[str, str], here: str) -> Optional[Tuple[str, str]]:
    """``(blank, entered)`` - Excel's two answers to "is any entry of
    period ``i`` filled?", or None when a row is not laid out on the
    line's months or the test would not fit in one Excel formula.
    COUNTA, not COUNT: an entered error or TRUE/FALSE is an entry too."""
    cells = []
    for row in rows:
        addr = addresses.get(row)
        if addr is None or len(addr.values) != n:
            return None
        cells.append(_qualified(addr.values[i], var_to_sheet.get(row), here))
    if not cells:
        return None
    if len(cells) == 1:
        return f'{cells[0]}=""', f'{cells[0]}<>""'
    count = "+".join(
        f"COUNTA({','.join(cells[k:k + _MAX_ARGS])})"
        for k in range(0, len(cells), _MAX_ARGS)
    )
    if len(count) > _MAX_COUNT_CHARS:
        return None
    return f"{count}=0", f"{count}>0"


def _cell(addr: Optional[VariableAddresses], i: int) -> Optional[str]:
    if addr is None or not addr.values:
        return None
    if len(addr.values) == 1:
        return addr.values[0]          # a scalar broadcasts
    return addr.values[i] if i < len(addr.values) else None


def write_blend_links(
    ws: Any,
    sheet_mv: MultiVariableBase,
    addresses: Dict[str, VariableAddresses],
    var_to_sheet: Dict[str, str],
    sheet_name: str,
) -> None:
    """Stamp the splice formulas onto this sheet's blend rows."""
    from ...core.tracks_decl import resolve_tracks_decl

    decl = resolve_tracks_decl(sheet_mv)
    spec = getattr(decl, "blend", None) if decl is not None else None
    if spec is None:
        return

    by_id = {
        comp.id: comp
        for _, comp in sheet_mv.walk()
        if isinstance(comp, Variable) and comp.id
    }
    tracked = {vid[: vid.index(TRACK_INFIX)] for vid in addresses if TRACK_INFIX in vid}
    for comp in by_id.values():
        if TRACK_INFIX in comp.id:
            continue                      # a subrow, not the blend row
        if not getattr(comp, "_blend_is_splice", False):
            continue                      # derived: its own formula stands
        if var_to_sheet.get(comp.id) != sheet_name:
            continue
        addr = addresses.get(comp.id)
        if addr is None or not addr.values:
            continue
        given = addresses.get(f"{comp.id}{TRACK_INFIX}{spec.given}")
        follow = addresses.get(f"{comp.id}{TRACK_INFIX}{spec.follow}")
        if given is None and follow is None:
            continue                      # nothing to reference
        n = len(addr.values)
        entries = {
            role: _side_entries(comp, by_id.get(f"{comp.id}{TRACK_INFIX}{role}"),
                                role, addresses, tracked)
            for role in (spec.given, spec.follow)
        }

        def entered(role: str, i: int) -> Optional[Tuple[str, str]]:
            rows = entries[role]
            return (_entered(rows, i, n, addresses, var_to_sheet, sheet_name)
                    if rows else None)

        splice = _before_boundary(sheet_mv, spec, n)
        for i, ref in enumerate(addr.values):
            g = _cell(given, i)
            f = _cell(follow, i)
            first, second = (g, f) if splice[i] else (f, g)
            first_role = spec.given if splice[i] else spec.follow
            if first and second:
                test = entered(first_role, i)
                blank = test[0] if test else f'{first}=""'
                ws[ref] = f"=IF({blank},{second},{first})"
            elif first:
                ws[ref] = f"={first}"
            elif second:
                ws[ref] = f"={second}"
            if g:
                test = entered(spec.given, i)
                _mark_settled(ws, ref, test[1] if test else f'{g}<>""')


#: The blue font that marks a blended actual — the finance-Excel input blue.
SETTLED_FONT_COLOR = "FF2563EB"


def _mark_settled(ws: Any, ref: str, entered: str) -> None:
    """Ink the blended cell blue WHILE an actual stands behind it.

    A conditional rule, not a stamped font: the cell itself is a live
    splice formula, so a baked colour would freeze the provenance the
    moment the file is written — the same dead-values habit this
    module exists to kill. Typed into the file, an actual turns its own
    month blue on the spot.

    Font only, never a fill: a fill would win over the house style's
    section bands and cut holes in them.
    """
    try:
        from openpyxl.formatting.rule import FormulaRule
        from openpyxl.styles import Font
    except ImportError:      # pragma: no cover - openpyxl is a hard dep
        return
    ws.conditional_formatting.add(
        ref,
        FormulaRule(formula=[entered],
                    font=Font(color=SETTLED_FONT_COLOR)),
    )
