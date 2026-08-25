# SPDX-License-Identifier: Apache-2.0
"""Track subrows, written as FORMULAS — parity reaches every track.

A tracked line emits one row per track (``expand_tracked_tree``). The
blend row has always carried the line's formula, rendered against its
operands' blend rows — the live-universe law. The other subrows got
their numbers and nothing else, so «Итого доходы · план» read as a
frozen literal while the very same line one row up explained itself.
Change January's план in the downloaded file and the план total sat
still: the dead-values habit Excel parity exists to kill, surviving in
the one place nobody looked.

The fix needs no new arithmetic, because the engine already computed
the answer: a DERIVED track's value IS the line's own expression over
its operands' same-track values. So the same expression, translated
through an address book where every tracked operand resolves to ITS
subrow of that track, is the honest formula for the cell — the exact
indirection ``totals_writer.bucket_cells`` uses to spell a bucket over
bucket cells.

What keeps its number, by construction:

* an AUTHORED track (``mo.Variable(..., факт=[...])``) — that series is
  DATA, and data has no formula, exactly like an input cell;
* a line with no expression at all — same reason;
* a track whose operand does not carry it, so the reference would have
  to borrow another track's row and the formula would contradict the
  number it sits on;
* the blend row itself, which ``blend_writer`` owns.
"""

from typing import Any, Dict, Optional

from .addresses import VariableAddresses
from .blend_writer import TRACK_INFIX
from .translator import ExcelTranslator
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable


def _role_book(
    addresses: Dict[str, VariableAddresses], role: str
) -> Dict[str, VariableAddresses]:
    """The address book as seen FROM one track.

    Every line that has a subrow for ``role`` answers with that
    subrow's cells; everything else — constants, untracked lines,
    cross-sheet operands — keeps its own address, which is what a
    hand-built book would reference too.
    """
    book = dict(addresses)
    suffix = f"{TRACK_INFIX}{role}"
    for vid, addr in addresses.items():
        if vid.endswith(suffix):
            book[vid[: -len(suffix)]] = addr
    return book


def _role_of(vid: str) -> Optional[str]:
    head, sep, role = vid.partition(TRACK_INFIX)
    return role if sep and role else None


def write_track_formulas(
    ws: Any,
    sheet_mv: MultiVariableBase,
    addresses: Dict[str, VariableAddresses],
    var_to_sheet: Dict[str, str],
    sheet_name: str,
) -> None:
    """Stamp each DERIVED track subrow with the line's own formula,
    read in that track's coordinates."""
    tracked_ids = {
        vid[: vid.index(TRACK_INFIX)]
        for vid in addresses
        if TRACK_INFIX in vid
    }
    # Same identity homes as the main writer: a subrow formula's
    # loop-built operands resolve to the subrow cells that hold them
    # (the ОПВ-cap-repeated-four-times case), never unrolled. Built
    # PER ROLE below: a shared object lives in several subrows (an
    # inherited прогноз holds the план's very elements), and this
    # role's own subrow must win — a прогноз formula referencing a
    # план cell would sit still when the file edits the прогноз row.
    from .translator import identity_cells_for
    from .writer import _identity_homes
    generic_homes = _identity_homes([sheet_mv], addresses)
    by_role_homes: Dict[str, Dict[int, "tuple[str, int]"]] = {}

    def _homes_for(role: str) -> Dict[int, "tuple[str, int]"]:
        cached = by_role_homes.get(role)
        if cached is None:
            suffix = f"{TRACK_INFIX}{role}"
            # The wrapper expressions live on the HEAD's ListExpr (a
            # subrow is a flat value copy), but the cells that hold the
            # series are the ROLE's subrow — so the builder reads the
            # head and targets the subrow. Derived roles only: an
            # authored role's subrow is data, and pointing formulas at
            # data is fine, but pointing THROUGH a head whose positional
            # series the authored role does not share would lie.
            role_rows = []
            for vid, comp in by_id.items():
                if not vid.endswith(suffix):
                    continue
                head = by_id.get(vid[: -len(suffix)])
                if head is not None and getattr(
                        comp, '_authored_track', None) != role:
                    role_rows.append((vid, head))
            cached = {**generic_homes,
                      **identity_cells_for(role_rows, addresses)}
            by_role_homes[role] = cached
        return cached
    by_id = {
        comp.id: comp
        for _, comp in sheet_mv.walk()
        if isinstance(comp, Variable) and comp.id
    }

    for vid, comp in by_id.items():
        role = _role_of(vid)
        if role is None:
            continue                       # a blend row — not ours
        if var_to_sheet.get(vid) != sheet_name:
            continue
        if getattr(comp, "_authored_track", None) == role:
            continue                       # entered, not derived
        expr = getattr(comp, "_expr", None)
        if expr is None:
            # The subrow is a flat copy; the EXPRESSION lives on the
            # line's blend row, which shares the same head id.
            # ``is not None``, never truthiness: a list Variable
            # refuses ``__bool__`` on purpose.
            owner = by_id.get(vid[: vid.index(TRACK_INFIX)])
            if owner is None:
                continue
            expr = getattr(owner, "_expr", None)
            if expr is None:
                continue
            deps = owner.dependencies
        else:
            deps = comp.dependencies
        # Every tracked operand must carry THIS track, or the formula
        # would quietly read a different one.
        if any(
            dep in tracked_ids
            and f"{dep}{TRACK_INFIX}{role}" not in addresses
            for dep in deps
        ):
            continue
        addr = addresses.get(vid)
        if addr is None or not addr.values:
            continue
        book = _role_book(addresses, role)
        translator = ExcelTranslator(book, var_to_sheet,
                                     identity_cells=_homes_for(role))
        for i, ref in enumerate(addr.values):
            # Empty periods keep the formula too — Excel's own rule,
            # and the founder's call: «пусто + пусто» is 0 there, and a
            # live formula that answers 0 beats a frozen blank. The one
            # place this could speak louder than the engine is a SPLICE
            # line, whose blend row asks ``=IF(факт="",…)`` — but a
            # spliced track is AUTHORED, and authored tracks never
            # reach this loop.
            try:
                rendered = translator.translate(
                    expr, i, current_sheet=sheet_name,
                    self_address=addr, self_var=comp,
                )
            except Exception:
                # Presentation only — a cell that will not translate
                # keeps the number the engine already wrote there.
                break
            # The same refusal the main writer makes: a translation
            # that collapsed a foreign list positionally would compute
            # the WRONG number, and a wrong formula is worse than an
            # honest literal. The truth wins — keep the value.
            if getattr(translator, "last_render_lossy", False):
                continue
            ws[ref] = rendered
