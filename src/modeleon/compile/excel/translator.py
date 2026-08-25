# SPDX-License-Identifier: Apache-2.0
"""ExcelTranslator — Excel-specific facade over the IR walker.

The interesting work lives in two layers underneath:

- :class:`~modeleon.compile.walker.Walker` — generic AST dispatcher.
- :class:`~modeleon.compile.excel_backend.ExcelRenderer` — one concrete
  renderer. Implements
  :class:`~modeleon.compile.renderer.Renderer` by rendering each AST
  node as an Excel formula fragment.

``ExcelTranslator`` just wires those two together and adds the
Excel-specific postprocessing — ``=`` prefix and outer-paren stripping
— that callers like ``writer.py`` and ``display/html.py`` already rely
on. The public ``.translate(expr, period_idx, current_sheet,
self_address)`` API is unchanged.

Adding a second renderer (JSON, Google Sheets, pandas, SQL) is a matter
of writing a new ``FooBackend`` alongside ``ExcelRenderer`` and a
matching facade — no change to the walker, the AST, or any core code.
"""

from __future__ import annotations

from typing import Any, Dict, Iterable, Optional

from .addresses import VariableAddresses
from ...core.expr import Expr
from ..renderer import RenderCtx
from .renderer import ExcelRenderer, _strip_outer_parens
from ..walker import Walker


class ExcelTranslator:
    """Translate Modeleon ``Expr`` trees to Excel formula strings.

    Thin facade: constructs an :class:`ExcelRenderer` and its
    :class:`Walker`, then delegates ``.translate()`` to them while
    adding the ``=``-prefix / outer-paren cleanup Excel expects.
    """

    def __init__(
        self,
        addresses: Dict[str, VariableAddresses],
        var_to_sheet: Optional[Dict[str, str]] = None,
        identity_cells: Optional[Dict[int, "tuple[str, int]"]] = None,
    ):
        self.renderer = ExcelRenderer(addresses, var_to_sheet,
                                      identity_cells=identity_cells)
        self.walker = Walker(self.renderer)

    # Back-compat shims — some callers read these through the
    # translator instance rather than holding their own copies.
    @property
    def addresses(self) -> Dict[str, VariableAddresses]:
        return self.renderer.addresses

    @property
    def var_to_sheet(self) -> Dict[str, str]:
        return self.renderer.var_to_sheet

    def translate(
        self,
        expr: Expr,
        period_idx: int = 0,
        current_sheet: Optional[str] = None,
        self_address: Optional[VariableAddresses] = None,
        self_var: Optional[Any] = None,
    ) -> str:
        """Render ``expr`` as an Excel formula string (prefixed with ``=``).

        Args:
            expr: Expression tree. Must be an :class:`Expr` subclass;
                there is no raw-string fallback.
            period_idx: Column index (0-based) within the result
                Variable's cell range.
            current_sheet: Sheet the result is being written to, used
                to suppress redundant sheet qualifiers when the
                referenced Variable lives on the same sheet.
            self_address: Address object for the Variable being
                translated; required for ``SelfRef`` nodes with
                ``period_idx > 0`` so ``{prev}`` can resolve to the
                previous period's cell.
            self_var: Owning Variable of ``expr``. Used by fallback
                paths (backend-specific ``FuncCall`` nodes) to inline
                ``self_var.value`` when the function isn't natively
                renderable to Excel.
        """
        if expr is None:
            return ""
        if not isinstance(expr, Expr):
            raise TypeError(
                f"ExcelTranslator.translate() expects an Expr node, got "
                f"{type(expr).__name__}"
            )
        ctx = RenderCtx(
            period_idx=period_idx,
            current_sheet=current_sheet,
            self_address=self_address,
            self_var=self_var,
        )
        emitted = self.walker.render(expr, ctx)
        # A lossy positional collapse (see RenderCtx.lossy_inline) —
        # surfaced to the caller so the cell can carry the VALUE
        # instead of a formula that computes something else.
        self.last_render_lossy = bool(ctx.lossy_inline)
        # recurrence templates whose {prev} slot is a parenthesized start
        # value (``(x + 1) * (1 - churn)``) leave redundant outer
        # parens; drop them before prefixing ``=`` so the cell reads
        # ``=x + 1`` not ``=(x + 1)``.
        emitted = _strip_outer_parens(emitted)
        return emitted if emitted.startswith("=") else f"={emitted}"


def identity_cells_for(
    rows: "Iterable[tuple[str, Any]]",
    addresses: Dict[str, VariableAddresses],
) -> Dict[int, "tuple[str, int]"]:
    """``id(floating month element) → (owner vid, period)``.

    A loop-built model computes intermediates as per-month expressions
    in local Python lists, then materializes the SAME list as a row —
    so the row's month cell and the operand inside a downstream formula
    are one object. Nothing else connects them: the operand is
    address-less by id, and rendering it unrolled the whole expression
    into every consumer (a payroll cell repeating the ОПВ cap four
    times, one row under the ОПВ row that holds it).

    Identity is the honest join. Only address-less elements enter the
    map (a rooted row resolves by id already), and the FIRST home wins:
    walk order puts a row before the rows that read it, so an element
    shared across rows references its defining row, deterministically.

    ``rows`` decides policy: the caller chooses which rows may serve as
    homes (an expanded tracked tree must offer SUBROWS, never the blend
    row — a план formula referencing a blend cell would move when a
    fact lands).
    """
    from ...core.expr import ListExpr, VarRef

    out: Dict[int, "tuple[str, int]"] = {}

    def _home(key_obj: Any, vid: str, period: int) -> None:
        if key_obj is not None:
            out.setdefault(id(key_obj), (vid, period))

    for vid, var in rows:
        addr = addresses.get(vid)
        if addr is None or not addr.values:
            continue
        # The join keys live in the row's EXPRESSION, not its value:
        # ``mo.Variable([месячные выражения])`` evaluates the elements
        # into plain numbers for the value list, wraps each into a
        # fresh Variable for its ListExpr — but the wrapper SHARES the
        # original's expression tree, and downstream formulas reach
        # that same tree through their own operands. So a consumer's
        # operand is recognized either as the wrapper itself or by the
        # expression object they share.
        expr = getattr(var, '_expr', None)
        if isinstance(expr, ListExpr):
            for i, item in enumerate(getattr(expr, 'items', ())):
                if i >= len(addr.values):
                    break
                if isinstance(item, VarRef):
                    _home(item.var, vid, i)
                    _home(getattr(item.var, '_expr', None), vid, i)
        # Rows whose value list still holds live Variables (authored
        # lists of records' cells, splice series) join directly.
        value = getattr(var, '_value', None)
        if isinstance(value, list):
            for i, el in enumerate(value):
                if i >= len(addr.values):
                    break
                el_id = getattr(el, 'id', None)
                if el_id is None or el_id in addresses:
                    continue
                _home(el, vid, i)
                _home(getattr(el, '_expr', None), vid, i)
    return out
