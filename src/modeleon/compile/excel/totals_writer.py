# SPDX-License-Identifier: Apache-2.0
"""Writing the subtotal cells — the CONTENT half of ``totals``.

Geometry (which column each bucket owns) is the layout's
:class:`~.totals.TotalsPlan`; this module fills the cells so they are
ALIVE:

* rule lines — ``=SUM(янв,фев,мар)`` / ``=AVERAGE(...)`` / the last
  month's ref, straight from the line's regrain rule;
* formula lines — the line's own formula rendered AT the bucket, via a
  translator whose address book maps every periodic line to its bucket
  cells (:func:`project_model` supplies the bucket-grain expression
  tree, and it keys lines by the same qpaths);
* ruleless literals — the projection's ``#VALUE!`` teaching token.

Edit January in the downloaded file and the quarter and year move —
the parity promise, kept in the totals too.
"""

import logging
from typing import Any, Dict

from .addresses import VariableAddresses
from .translator import ExcelTranslator
from .totals import TotalsPlan, regrain_rule_of, rule_formula
from .writer import resolve_number_format
from ...core.multi_variable import MultiVariableBase
from ...core.variable import Variable

logger = logging.getLogger(__name__)


def _walk_vars(mv: MultiVariableBase) -> Dict[str, Variable]:
    out: Dict[str, Variable] = {}
    for _, comp in mv.walk():
        if isinstance(comp, Variable) and comp.id:
            out[comp.id] = comp
    return out


def bucket_cells(
    sheet_mv: MultiVariableBase,
    plan: TotalsPlan,
    addresses: Dict[str, VariableAddresses],
    var_to_sheet: Dict[str, str],
    sheet_name: str,
    projected: Dict[str, Dict[str, Variable]] | None = None,
) -> Dict[str, Dict[str, Any]]:
    """Every bucket cell's CONTENT — ``{ref: {"formula": str|None,
    "value": Any, "id": vid}}`` — the single computation behind both
    the .xlsx writer and the run-state wire, so the file and the
    grid's formula mirror cannot disagree.

    ``projected`` — pre-computed ``{kind: {rel_id: projected Var}}``
    from a caller that already ran :func:`project_model` (the wire
    memoizes one projection per kind across sheets); ``None`` projects
    here, which is fine for the writer's one-shot export."""
    from ...core.projection import project_model

    native = _walk_vars(sheet_mv)
    # The projected tree roots at the SHEET, the native ids at the
    # MODEL — «m.план.налог» vs «план.налог». Key both by the path
    # RELATIVE to their roots so the same line meets itself.
    sheet_prefix = f"{sheet_mv.id}." if sheet_mv.id else ""

    def _rel(vid: str) -> str:
        return vid[len(sheet_prefix):] if vid.startswith(sheet_prefix) else vid

    if projected is None:
        projected = {}
        for kind in plan.kinds:
            try:
                ptree = project_model(sheet_mv, kind, on_error="absorb")
                pprefix = f"{ptree.id}." if ptree.id else ""
                projected[kind] = {
                    (pid[len(pprefix):] if pid.startswith(pprefix) else pid): v
                    for pid, v in _walk_vars(ptree).items()
                }
            except Exception as exc:  # projection must never kill the book
                logger.warning("totals: %s projection failed: %s", kind, exc)
                projected[kind] = {}

    # One translator per kind: every periodic line's "periods" are its
    # BUCKET cells, everything else (constants, cross-sheet operands)
    # keeps its native address — which is exactly how a hand-built
    # book's quarter formula reads.
    translators: Dict[str, ExcelTranslator] = {}
    bucket_refs: Dict[str, Dict[str, list]] = {}
    for kind in plan.kinds:
        merged = dict(addresses)
        sheets = dict(var_to_sheet)
        refs_of: Dict[str, list] = {}
        for vid, addr in addresses.items():
            if var_to_sheet.get(vid) != sheet_name:
                continue
            if vid not in native:
                continue
            pvar = projected[kind].get(_rel(vid))
            if getattr(native[vid], '_metric_no_bucket', False):
                # «% роста» has no prior-year bucket yet — the grid
                # shows null there, and the file shows empty, never a
                # projected lag whose step no longer fits the grain.
                continue
            if _is_periodic(native[vid]):
                row = _row_of(addr.values[0])
                # The line's SLOT inside each group — a track column
                # totals into its own bucket subcolumn, exactly under
                # its months.
                _slot = getattr(native[vid], '_col_slot', 0) or 0
                refs = [
                    f"{_col_letter(plan.cell((kind, bi), _slot))}{row}"
                    for bi in range(len(plan.buckets[kind]))
                ]
                refs_of[vid] = refs
                slot = VariableAddresses(
                    name="", formula=refs[0], values=refs,
                )
            else:
                # A constant reads the same in every bucket — its
                # native cell IS its quarter and its year.
                slot = addr
            merged[vid] = slot
            if pvar is not None and pvar.id:
                # The projected expr's refs point at PROJECTED objects
                # («план.выручка»), the address book at native ids
                # («m.план.выручка») — alias, or every operand falls
                # back to an inlined literal and the bucket goes dead.
                merged[pvar.id] = slot
                sheets[pvar.id] = sheet_name
        translators[kind] = ExcelTranslator(merged, sheets)
        bucket_refs[kind] = refs_of

    out: Dict[str, Dict[str, Any]] = {}
    for vid, refs in _stable(bucket_refs, plan):
        kind = refs["kind"]
        var = native[vid]
        row_refs = bucket_refs[kind][vid]
        month_row = _row_of(addresses[vid].values[0])
        rule = regrain_rule_of(var)
        pvar = projected.get(kind, {}).get(_rel(vid))
        for bi, (_label, lo, hi) in enumerate(plan.buckets[kind]):
            formula: Any = None
            value: Any = None
            if rule is not None:
                _slot = getattr(var, '_col_slot', 0) or 0
                month_refs = [
                    f"{_col_letter(plan.cell(('month', i), _slot))}{month_row}"
                    for i in range(lo, min(hi, plan.n))
                    if i < len(addresses[vid].values)
                ]
                if month_refs:
                    formula = rule_formula(rule, month_refs)
            elif pvar is not None and getattr(pvar, "_expr", None) is not None:
                try:
                    formula = translators[kind].translate(
                        pvar._expr,
                        bi,
                        current_sheet=sheet_name,
                        self_address=row_refs[bi],
                        self_var=pvar,
                    )
                except Exception:
                    value = _bucket_value(pvar, bi)
            elif pvar is not None:
                # Ruleless literal → the projection already answered:
                # a value where it could, its teaching token where it
                # could not. Never an invented sum.
                value = _bucket_value(pvar, bi)
            out[row_refs[bi]] = {
                "formula": formula, "value": value, "id": vid,
            }
    return out


def write_totals(
    ws: Any,
    sheet_mv: MultiVariableBase,
    plan: TotalsPlan,
    addresses: Dict[str, VariableAddresses],
    var_to_sheet: Dict[str, str],
    sheet_name: str,
) -> None:
    from openpyxl.styles import Font

    native = _walk_vars(sheet_mv)
    cells = bucket_cells(sheet_mv, plan, addresses, var_to_sheet, sheet_name)
    for ref, spec in cells.items():
        cell = ws[ref]
        if spec["formula"] is not None:
            cell.value = spec["formula"]
        elif spec["value"] is not None:
            cell.value = spec["value"]
        cell.font = Font(bold=True)
        var = native.get(spec["id"])
        fmt = resolve_number_format(var) if var is not None else None
        if fmt:
            cell.number_format = fmt


def _stable(bucket_refs: Dict[str, Dict[str, list]], plan: TotalsPlan):
    for kind in plan.kinds:
        for vid in bucket_refs[kind]:
            yield vid, {"kind": kind}


def _bucket_value(pvar: Variable, bi: int) -> Any:
    v = getattr(pvar, "_value", None)
    if type(v).__name__ == "TrackValues":
        # A tracked line's projected value is the whole role→series
        # container — written raw it lands in the cell as machine
        # repr. Resolve to the display default (the blend when the
        # model blends, else the first declared track) — the same
        # law the grid row shows.
        from ...core.tracks_decl import resolve_tracks_decl

        decl = resolve_tracks_decl(getattr(pvar, "_owner", None))
        blend_name = getattr(getattr(decl, "blend", None), "name", None)
        roles = list(v.roles)
        v = v[blend_name if blend_name in roles else roles[0]]
    if isinstance(v, list):
        return v[bi] if bi < len(v) else None
    return v


def _is_periodic(var: Variable) -> bool:
    value = getattr(var, "_value", None)
    if isinstance(value, list) or type(value).__name__ == "TrackValues":
        return True
    return bool(getattr(var, "_indexed_by", ()) or ())


def _row_of(address: str) -> int:
    i = 0
    while i < len(address) and address[i].isalpha():
        i += 1
    return int(address[i:] or 0)


def _col_letter(col: int) -> str:
    result = ""
    while col > 0:
        col -= 1
        result = chr(65 + col % 26) + result
        col //= 26
    return result
