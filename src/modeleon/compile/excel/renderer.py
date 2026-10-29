# SPDX-License-Identifier: Apache-2.0
"""Excel renderer for the AST walker.

Implements :class:`~modeleon.compile.renderer.Renderer` by rendering each
AST node to an Excel formula fragment. The :class:`ExcelTranslator`
(in ``translator.py``) is the thin facade that wires this renderer up
to a :class:`~modeleon.compile.walker.Walker` and adds the
Excel-specific postprocessing (``=`` prefix, outer-paren stripping).

All Excel specifics live here: operator translation, literal
formatting, cell-reference syntax, range notation, sheet qualification,
function name casing. Another renderer (the JSON renderer in
``compile/json`` is one) is a parallel file alongside this one — core,
translator-facade, and walker stay untouched.
"""

from __future__ import annotations

import contextlib
import logging
import re
import warnings
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Callable, Dict, Iterable, Iterator, List, Optional

from .addresses import VariableAddresses
from ...core.errors import CrossScopeReferenceWarning
from ...core.expr import (
    BinOp,
    Compare,
    Expr,
    FuncCall,
    ListExpr,
    Literal,
    MethodCall,
    Paren,
    Regrain,
    Restrict,
    RollingAggregate,
    SelfRef,
    Subscript,
    TimeRef,
    UnaryOp,
    VarRef,
)
from ..renderer import RenderCtx
from ...functions.registry import (
    excel_name,
    file_prefix,
    is_passthrough,
    is_range_taking,
    per_period_arity,
    range_positions,
)

logger = logging.getLogger(__name__)


# Sheet names that are simple identifiers (starts with letter/underscore,
# rest word chars) can appear unquoted in a reference; anything else
# (spaces, dashes, punctuation, leading digit) needs single-quote wrapping.
# ``\Z``, not ``$``: ``$`` also matches before a trailing newline, and a
# sheet named 'Plan\n' would be written unquoted as ``Plan\n!B2``.
_UNQUOTED_SHEET_NAME = re.compile(r"[A-Za-z_]\w*\Z")

#: Excel keeps at most 8192 characters of one formula, the ``=`` among
#: them; a longer one sends the whole book to repair.
_MAX_FORMULA_CHARS = 8192

#: Excel's limit on the arguments of one call.
_MAX_ARGS = 255

# A cell reference on its own — an optional sheet, quoted or bare, and
# a coordinate: the only term a ``+`` chain may fold into a range.
_CELL_TERM = re.compile(r"(?:'(?:[^']|'')*'!|[A-Za-z_]\w*!)?[A-Z]{1,3}[1-9]\d*\Z")


def _fold_sum(parts: List[str]) -> Optional[str]:
    """``SUM`` over the runs of a ``+`` chain whose terms are all cells
    of one sheet, each run its consecutive cells in order — the way a
    hand-built book writes a long sum. None for any other chain: a
    minus in it, a term that is not a cell, cells of two sheets."""
    terms, ops = parts[0::2], parts[1::2]
    if any(op != "+" for op in ops) or not all(_CELL_TERM.fullmatch(t) for t in terms):
        return None
    prefixes = {t.rpartition("!")[0] for t in terms}
    if len(prefixes) != 1:
        return None
    from .totals import compact_refs
    prefix = prefixes.pop()
    prefix = f"{prefix}!" if prefix else ""
    runs = compact_refs([t.rpartition("!")[2] for t in terms])
    if len(runs) > 255:                # SUM takes at most 255 arguments
        return None
    return "SUM({})".format(", ".join(
        f"{prefix}{first}:{last}" if last else f"{prefix}{first}" for first, last in runs))


@dataclass(frozen=True)
class _Frame:
    """The value a formula is being written an element of, as Python
    computed it: ``var`` (the cell's own line, or an unnamed intermediate
    written inline) and the element's index less the period being written
    (``offset``: a lag that writes the previous period moves both).
    ``window`` is the window (a sheet window key) ``var``'s operands were
    read onto by DATE — an operator or a function that is one call per
    period aligning two windows — and None where they were read by
    position (or by key)."""

    var: Any
    offset: int
    window: Optional[tuple]


def _aligned_key(var: Any) -> Optional[tuple]:
    """The window a formula's operands were aligned onto by date — the
    combined window an operator or a one-call-per-period function puts its
    result on and marks (``_read_by_date``) — as a sheet window key; None
    for a value its operands gave by position (a line on a window of its
    own read otherwise: ``start=``, a re-grain, a copy)."""
    if not getattr(var, '_read_by_date', False):
        return None
    if getattr(var, '_grain', None) is None or getattr(var, '_start', None) is None:
        return None
    value = getattr(var, '_value', None)
    n = len(value) if isinstance(value, list) else getattr(value, 'time_length', None)
    if not n:
        return None
    first_day = getattr(var, '_first_day', None)
    key = (var._grain, var._start, n)
    return key if first_day is None else (*key, first_day)


def _by_position(var: Any) -> bool:
    """A value whose elements are read at the cell's own period: not
    aligned by date, along no keys — writing it needs no frame."""
    return not getattr(var, '_read_by_date', False) and not getattr(var, '_keys', None)


def _periods_from(a: tuple, b: tuple) -> int:
    """The periods from window ``a``'s first to window ``b``'s, both of one
    grain: negative when ``b`` starts earlier."""
    from ...core.time import _count, _first_day, _parse
    grain = a[0]
    here, there = _parse(a[1], grain), _parse(b[1], grain)
    if _first_day(here, grain) <= _first_day(there, grain):
        return _count(here, there, grain) - 1
    return 1 - _count(there, here, grain)


_PY_TO_EXCEL_OP: Dict[str, str] = {
    "+": "+",
    "-": "-",
    "*": "*",
    "/": "/",
    "**": "^",
    "==": "=",
    "!=": "<>",
    "<": "<",
    "<=": "<=",
    ">": ">",
    ">=": ">=",
}


# Operator precedence in Excel formulas. Higher binds tighter. A child BinOp
# whose precedence is *strictly less* than the parent's must be wrapped in
# parens; same-precedence children only need wrapping on the right side of
# non-commutative operators (``a - (b - c)`` ≠ ``(a - b) - c``).
_BINOP_PRECEDENCE: Dict[str, int] = {
    "**": 4, "^": 4,
    "*": 3, "/": 3, "//": 3, "%": 3,
    "+": 2, "-": 2,
    "&": 1,
    "==": 0, "!=": 0, "<": 0, "<=": 0, ">": 0, ">=": 0,
}

# Non-commutative-on-the-right ops: same-precedence right child needs parens.
# `a - b - c` is fine (left-associative) but `a - (b - c)` must keep them.
_NON_COMMUTATIVE_RIGHT: set[str] = {"-", "/", "//", "%", "**", "^"}

# Operators of one precedence that read left to right: a left operand of
# the same level never takes parens, so a run of them renders flat.
_CHAIN_LEVELS = (frozenset({"+", "-"}), frozenset({"*", "/"}))


# Functions that accept an Excel range for any list-valued VarRef argument.
# The translator renders those args as ranges (``B2:F2``) and leaves
# scalar VarRefs / literals as normal. Covers classic aggregates (SUM,
# MAX, MIN, …) and financial-series funcs (IRR, NPV, XIRR) where Excel
# expects a range for the cash-flow / dates argument. A call over enough
# values with a series among them is the exception (``_per_period_call``):
# one call per period, its operands this period's cells — ``SUM(a, b)``,
# ``MIN`` / ``MAX`` over two or more operands, ``MEDIAN(a, b)``, a
# percentile over a list of yearly lines.
def _contiguous_run(refs: list) -> bool:
    """Do these cell refs sit side by side on one row, or one under
    another in one column, in order?

    A range is only honest when they do. The interleave that breaks it
    is ordinary — a declared quarter total puts a column between March
    and April — and the resulting ``SUM`` looks perfectly valid while
    counting the quarter twice.
    """
    from openpyxl.utils.cell import coordinate_to_tuple

    try:
        cells = [coordinate_to_tuple(r.split("!")[-1].replace("$", ""))
                 for r in refs]
    except Exception:
        return True                    # unparseable: emit the plain range
    rows = {r for r, _c in cells}
    cols = {c for _r, c in cells}
    if len(rows) == 1:
        run = [c for _r, c in cells]
    elif len(cols) == 1:
        run = [r for r, _c in cells]
    else:
        return False
    return all(b - a == 1 for a, b in zip(run, run[1:]))


def _run_spans(refs: List[str]) -> List[tuple[int, int]]:
    """The runs of these cell refs, as ``[lo, hi)`` positions: each the
    longest stretch of cells side by side on one row, or one under another
    in one column, in order — one direction a run (``totals.compact_refs``
    walks them the same way). Total columns between a line's months break
    its row into such runs. Refs it cannot read are one run, as
    :func:`_contiguous_run` takes them."""
    from openpyxl.utils.cell import coordinate_to_tuple

    try:
        cells = [coordinate_to_tuple(r.split("!")[-1].replace("$", "")) for r in refs]
    except Exception:
        return [(0, len(refs))] if refs else []
    spans: List[tuple[int, int]] = []
    lo = 0
    step: Optional[tuple[int, int]] = None
    for i in range(1, len(cells)):
        d = (cells[i][0] - cells[i - 1][0], cells[i][1] - cells[i - 1][1])
        if d in ((0, 1), (1, 0)) and step in (None, d):
            step = d
            continue
        spans.append((lo, i))
        lo, step = i, None
    if cells:
        spans.append((lo, len(cells)))
    return spans


def _united_spans(
    rows: List[List[str]],
    runs: Callable[[List[str]], List[tuple[int, int]]] = _run_spans,
) -> List[tuple[int, int]]:
    """Spans of positions over which each of ``rows`` — ranges read side by
    side, cell ``i`` of each with cell ``i`` of the others — is one run:
    every row cut where any of them breaks. By position, never by column: a
    sheet with a start of its own puts its total columns after other
    periods. ``runs`` splits one row (a renderer passes the split it
    remembers)."""
    n = len(rows[0])
    cuts = sorted({lo for cells in rows for lo, _hi in runs(cells)} | {n})
    return list(zip(cuts, cuts[1:]))


#: The criteria that a number or a date meets and a blank, a text, a
#: boolean or an error does not: what an average counts.
_NUMBERS = '">=-9.99999999999999E+307"'


def _criteria_ranges(func: str, n: int) -> Optional[tuple[int, ...]]:
    """The positions of the ranges of a conditional count, sum or mean of
    ``n`` arguments — the others are its criteria — or None for any other
    function. These add up over the runs of their ranges."""
    if func == "COUNTIF":
        return (0,)
    if func in ("SUMIF", "AVERAGEIF"):
        return (0, 2) if n > 2 else (0,)
    if func == "COUNTIFS":
        return tuple(range(0, n, 2))
    if func in ("SUMIFS", "AVERAGEIFS"):
        return (0, *range(1, n, 2))
    return None


@dataclass(frozen=True)
class NotOneRun:
    """Why a call over rows is written as the cell's computed value: the
    cells it reads as one range are not one run (``limit`` None), or,
    written as their runs, the call passes a limit of the spreadsheet —
    ``'arguments'`` (``size`` arguments, past 255) or ``'length'`` (a
    formula of ``size`` characters, past 8,192). ``lines`` holds what it
    reads whose cells are not one run: the sheet, the line's id (None for
    a list of lines) and the cells, in order."""

    limit: Optional[str]
    lines: tuple[tuple[Optional[str], Optional[str], tuple[str, ...]], ...]
    size: int = 0


# Which names are written as ranges, which pass through as ``=NAME(args)``,
# and what the file stores before a name (``_xlfn.DAYS``) is the function
# registry's to say (``modeleon.functions.registry``); each function
# registers itself from its own module.


def _per_period_call(call: FuncCall,
                     lists_series: Optional[Callable[[ListExpr], bool]] = None) -> bool:
    """A call over enough values to be one call per period, a series among
    them (the registry's ``per_period_arity``) — ``SUM(a, b)`` is a
    subtotal of a period's cells, ``MIN`` / ``MAX`` over two or more
    operands are a cap or a floor, ``MEDIAN(a, b)`` the median of a
    period's two cells, ``PERCENTILE_INC([a, b, c], k)`` the percentile of
    a period's cells of three lines: its values are written as this
    period's cells, never as the lines' ranges. The values are the
    arguments read as ranges, each line a list of lines holds counted;
    ``lists_series`` answers whether a list of lines holds a series."""
    func = call.func.upper()
    arity = per_period_arity(func)
    if arity is None:
        return False
    positions = range_positions(func)
    values: List[Any] = []
    lists: List[ListExpr] = []
    count = 0
    for i, arg in enumerate(call.args):
        if positions is not None and i not in positions:
            continue
        if isinstance(arg, ListExpr):
            lists.append(arg)
            count += len(arg.items)
        else:
            values.append(arg)
            count += 1
    if count < arity:
        return False
    if any(isinstance(arg, Literal) and isinstance(arg.value, list) for arg in values):
        return True
    series = lists_series or _lists_series
    return any(_is_series_ref(arg) for arg in values) or any(series(arg) for arg in lists)


def _lists_series(node: ListExpr) -> bool:
    return any(_is_series_ref(item) for item in node.items)


def _is_series_ref(node: Any) -> bool:
    """A reference to a line holding a series (a row, or a row per track)."""
    if not isinstance(node, VarRef):
        return False
    value = getattr(node.var, "_value", None)
    return isinstance(value, list) or hasattr(value, "roles")



def _value_at_period(value: Any, period_idx: int) -> Any:
    """Pick the period-th element from a list-valued Variable, or return
    the value unchanged when it's a scalar. Out-of-range indices clamp
    to the last element — matches how the layout engine handles
    length-1 broadcast."""
    if isinstance(value, list):
        if not value:
            return None
        return value[period_idx] if 0 <= period_idx < len(value) else value[-1]
    return value


def _value_to_excel_literal(value: Any) -> str:
    """Render a Python value as an Excel-literal string fit for embedding
    in a formula. Strings get double-quoted; booleans render as
    ``TRUE``/``FALSE``; None becomes ``0`` (matches how ``Literal(None)``
    is emitted elsewhere); dates become ``DATE(y, m, d)`` (a bare
    ``2025-01-01`` would be read by Excel as arithmetic); durations
    become their day count; numbers pass through via ``str()``."""
    if value is None:
        return "0"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        if value.hour or value.minute or value.second:
            # An instant keeps its clock: 13:00 is not the day's start. A
            # sum, it stands in brackets: ``x - start`` must not read
            # ``(x - DATE(…)) + TIME(…)``.
            return (f"(DATE({value.year}, {value.month}, {value.day})"
                    f"+TIME({value.hour}, {value.minute}, {value.second}))")
        value = value.date()
    if isinstance(value, date):
        return f"DATE({value.year}, {value.month}, {value.day})"
    if isinstance(value, timedelta):
        # A spreadsheet date is a count of days, so a duration is one
        # too: ``timedelta(days=90)`` → ``90``, ``timedelta(hours=12)``
        # → ``0.5``. Python's own ``90 days, 0:00:00`` is unreadable there.
        days = value / timedelta(days=1)
        return str(int(days)) if days.is_integer() else repr(days)
    if isinstance(value, str):
        escaped = value.replace('"', '""')
        return f'"{escaped}"'
    return str(value)


def _children(expr: Any) -> List[Expr]:
    """The sub-expressions of an AST node, in any node shape."""
    out: List[Expr] = []
    for name in ('inner', 'left', 'right', 'operand', 'base'):
        child = getattr(expr, name, None)
        if isinstance(child, Expr):
            out.append(child)
    for many in ('items', 'args'):
        out.extend(c for c in getattr(expr, many, None) or () if isinstance(c, Expr))
    kwargs = getattr(expr, 'kwargs', None) or {}
    out.extend(v for v in kwargs.values() if isinstance(v, Expr))
    return out


def _warn_out_of_scope_ref(var: Any, ctx: RenderCtx) -> None:
    """Fire :class:`CrossScopeReferenceWarning` when a formula's VarRef
    points at a Variable that isn't in this emission's layout.

    Two common causes (both user mistakes — see the warning class
    docstring): cross-model reference, or a floating Variable created
    outside any ``with MultiVariable(...):`` block. The warning names
    the orphan Variable so the fix is obvious.
    """
    label = getattr(var, "python_name", None) or getattr(var, "path", None)
    target_sheet = ctx.current_sheet or "the current emission"
    warnings.warn(
        f"Variable {label!r} is referenced from {target_sheet} but has no "
        f"cell address in this workbook — inlining its value as a literal. "
        f"Likely cause: the Variable was not attached to the model's MV tree "
        f"(via `parent.child = ...`), or belongs to a different model than "
        f"the one being emitted. Attach it under the model (or move the "
        f"formula to the model that owns it) to keep the reference live.",
        category=CrossScopeReferenceWarning,
        stacklevel=3,
    )


def _warn_unlaid_aggregate(func: str, arg: Expr, ctx: RenderCtx) -> None:
    """Fire :class:`CrossScopeReferenceWarning` when ``SUM`` / ``NPV`` / …
    reduces a list that has no cells in this workbook, so its computed
    value is written in place of the call."""
    label = "a literal list"
    if isinstance(arg, VarRef):
        name = (getattr(arg.var, "python_name", None)
                or getattr(arg.var, "_display_name", None))
        label = repr(name) if name else f"an unnamed list ({arg.var.path})"
    target_sheet = ctx.current_sheet or "the current emission"
    warnings.warn(
        f"{func} in {target_sheet} reduces {label}, which has no cells in "
        f"this workbook — writing the {func}'s computed value in its place "
        f"instead of a live {func}. Attach the list to the model "
        f"(`parent.child = ...`) to keep the {func} live.",
        category=CrossScopeReferenceWarning,
        stacklevel=3,
    )


def _has_top_level_operator(text: str) -> bool:
    """Whether a rendered fragment holds an operator outside its brackets
    and quotes - an expression, not a single cell, number or call."""
    depth = 0
    quote = None
    for i, ch in enumerate(text):
        if quote:
            if ch == quote:
                quote = None
        elif ch in "'\"":
            quote = ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif depth == 0 and ch in "+-*/^&=<>" and not (ch in "+-" and i == 0):
            return True
    return False


def _strip_outer_parens(s: str) -> str:
    """Remove redundant parens fully enclosing the string."""
    while len(s) >= 2 and s[0] == "(" and s[-1] == ")":
        depth = 0
        balanced = True
        for i, ch in enumerate(s):
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            if depth == 0 and i < len(s) - 1:
                balanced = False
                break
        if balanced:
            s = s[1:-1]
        else:
            break
    return s


class ExcelRenderer:
    """Walks an ``Expr`` tree and renders Excel formula fragments.

    Constructed with an addresses map (variable-id → cell addresses)
    and an optional var-to-sheet map for cross-sheet qualification.
    The :class:`Walker` wires itself onto ``self.walker`` at its own
    construction — backends use it to recurse into child nodes via
    ``self.walker.render(child, ctx)``.
    """

    def __init__(
        self,
        addresses: Dict[str, VariableAddresses],
        var_to_sheet: Optional[Dict[str, str]] = None,
        identity_cells: Optional[Dict[int, "tuple[str, int]"]] = None,
        time: Optional[Any] = None,
        periods: Optional[Dict[str, VariableAddresses]] = None,
    ) -> None:
        self.addresses = addresses
        #: Where a line read as a range stands, when ``addresses`` puts it
        #: elsewhere: a total column's formula reads a line at its bucket's
        #: cell, and the whole line — a range — at its own periods.
        self.periods = periods or {}
        #: The workbook's time (:class:`~.header.BookTime`): where a
        #: ``mo.time`` read lands — this sheet's header row, the anchor,
        #: or the start cell. ``None`` spells time off each row's window.
        self.time = time
        self.var_to_sheet = var_to_sheet or {}
        #: ``id(floating Variable) → (owner vid, period)`` — a per-month
        #: intermediate built in a Python loop is often the SAME object
        #: that sits in some laid-out row's value list. Without this map
        #: a formula over such intermediates unrolls them wholesale
        #: (e.g. a cell spelling a monthly subtotal out again instead of
        #: pointing at the row that already holds it);
        #: with it the reference resolves to the cell that already
        #: carries the number — the way a human builds the same sheet.
        self.identity_cells = identity_cells or {}
        #: What the formula being rendered wrote as a number: ``(node,
        #: literal)`` for each sub-expression inlined into a live formula,
        #: and the nodes no formula could stand for (the cell then
        #: carries its value), each with its cause where one is known
        #: (:class:`NotOneRun`). Per formula — the translator clears them
        #: before each walk.
        self.inlined: List[tuple[Expr, str]] = []
        self.lossy: List[Expr] = []
        self.lossy_causes: List[Optional[NotOneRun]] = []
        #: A list of lines many calls share (one range read once, see
        #: ``functions.criteria``) is read once per render: whether it
        #: holds a series, and the text of its cells per sheet and period
        #: (or why there is none) — keyed by the node, which each entry
        #: holds so its id stays its own. Nothing a render reads moves
        #: while it lasts.
        self._listed: Dict[int, tuple[ListExpr, bool]] = {}
        self._ranges: Dict[tuple[Any, ...],
                           tuple[ListExpr, Optional[str], Optional[NotOneRun]]] = {}
        #: The runs of a row's cells (``_run_spans``), by the cells: a
        #: row is split once, however many cells read it.
        self._spans: Dict[tuple[str, ...], List[tuple[int, int]]] = {}
        self.walker: Any = None  # set by Walker(renderer) constructor

    def _note_inlined(self, node: Expr, literal: str, ctx: RenderCtx) -> str:
        """``literal`` stands in the text for ``node`` — true now, stale
        once that operand's value moves (``RenderCtx.inlined_value``)."""
        ctx.inlined_value = True
        self.inlined.append((node, literal))
        return literal

    def _mark_lossy(self, node: Expr, ctx: RenderCtx,
                    cause: Optional[NotOneRun] = None) -> None:
        """No formula can stand for ``node`` here: the whole cell carries
        its computed value (``RenderCtx.lossy_inline``)."""
        ctx.lossy_inline = True
        self.lossy.append(node)
        self.lossy_causes.append(cause)

    def _ship_value(self, node: Expr, ctx: RenderCtx,
                    cause: Optional[NotOneRun] = None) -> str:
        """The cell's own computed value, in place of a formula that
        cannot be written around ``node``."""
        self._mark_lossy(node, ctx, cause)
        return _value_to_excel_literal(
            _value_at_period(getattr(ctx.self_var, "_value", None), ctx.period_idx))

    def _limited(self, node: Expr, text: str, args: int,
                 reads: Iterable[tuple[str, List[str]]], ctx: RenderCtx) -> str:
        """``text`` — a call over the rows ``reads`` (id and cells) written
        as their runs, of ``args`` arguments — or, past what one call or
        one formula of the spreadsheet holds, the cell's computed value,
        with the rows whose cells are not one run named."""
        if args > _MAX_ARGS:
            return self._ship_value(node, ctx, NotOneRun("arguments", self._split(reads), args))
        if len(text) >= _MAX_FORMULA_CHARS:
            return self._ship_value(
                node, ctx, NotOneRun("length", self._split(reads), len(text) + 1))
        return text

    # ─── Dispatch targets ───────────────────────────────────────

    def render_paren(self, node: Paren, ctx: RenderCtx) -> str:
        """Emit parens only when the wrapped subexpression is compound.

        ``Paren`` in the AST is a precedence-preservation hint, not a
        semantic requirement. When the inner node renders as an atomic
        token (cell reference, literal, function call), the parens are
        redundant — ``(C1) / C3`` carries no more meaning than ``C1 / C3``.
        """
        with self._chain_passes(node, node.inner, ctx):
            if self._needs_parens(node.inner):
                return f"({self.walker.render(node.inner, ctx)})"
            return self.walker.render(node.inner, ctx)

    def render_unaryop(self, node: UnaryOp, ctx: RenderCtx) -> str:
        return f"{node.op}{self.walker.render(node.operand, ctx)}"

    def render_compare(self, node: Compare, ctx: RenderCtx) -> str:
        op = _PY_TO_EXCEL_OP.get(node.op, node.op)
        return f"{self.walker.render(node.left, ctx)} {op} {self.walker.render(node.right, ctx)}"

    def render_listexpr(self, node: ListExpr, ctx: RenderCtx) -> str:
        """A ListExpr is positional: item *i* is the formula for period *i*.

        Render only the item the current cell asks for — the whole
        bracketed list is Python syntax, never a valid Excel expression.
        Out-of-range indices clamp to the last item, matching
        ``_value_at_period``.
        """
        if not node.items:
            raise ValueError(
                "ListExpr with no items has no per-period Excel rendering"
            )
        idx = ctx.period_idx if 0 <= ctx.period_idx < len(node.items) else len(node.items) - 1
        return self.walker.render(node.items[idx], ctx)

    def render_literal(self, node: Literal, ctx: RenderCtx) -> str:
        # ``x + [1, 2, 3]`` — a list operand is one item per period (an
        # aggregate over a literal list is written as its value before
        # its arguments ever render).
        return _value_to_excel_literal(_value_at_period(node.value, ctx.period_idx))

    def _placed(self, var_id: str, ctx: RenderCtx) -> str:
        """A row mirrored into this sheet's header is read from the
        mirror — except by the mirror itself, whose formula IS the
        reference to the row."""
        book = self.time
        if book is None:
            return var_id
        own = book.sheets.get(ctx.current_sheet)
        mirror = own.alias.get(var_id) if own is not None else None
        if mirror is None or getattr(ctx.self_var, 'id', None) == mirror:
            return var_id
        return mirror

    def render_timeref(self, node: TimeRef, ctx: RenderCtx) -> str:
        """A row whose formula IS time (``h.days = mo.time.days``)."""
        return self._render_time(ctx.self_var, node.field, ctx)

    def _render_time(self, host: Any, field: str, ctx: RenderCtx) -> str:
        """``mo.time.<field>`` at ``ctx.period_idx`` — ``host`` is the
        Variable that carries it (the time row itself, or the floating
        ``mo.time.<field>`` a formula read). Always a formula: the
        anchor's own chain, this sheet's header row, the anchor's row,
        or the period spelled off its window's start — and only through a
        header or anchor that counts the same window as the formula."""
        from ...core.time import (
            GRAINS, _first_day, _parse, resolve_default_window, spelled_first_day,
        )
        from .header import date_literal, spell_anchor, spell_inline, window_key
        i = ctx.period_idx
        self_var = ctx.self_var
        self_id = getattr(self_var, 'id', None)
        sheet = ctx.current_sheet
        window = resolve_default_window(self_var) if self_var is not None else None
        if window_key(window) is None:
            loc = getattr(host, 'time', None)
            if loc is None or loc.start is None or loc.grain not in GRAINS:
                return self._note_inlined(TimeRef(field), _value_to_excel_literal(
                    _value_at_period(getattr(host, '_value', None), i)), ctx)
            start = loc.first_day or _first_day(_parse(loc.start, loc.grain), loc.grain)
            return spell_inline(field, i, loc.grain, date_literal(start), loc.first_day)
        assert window is not None
        key = window_key(window)
        start0 = self._start_cell(window, sheet)
        own = self.time.sheets.get(sheet) if self.time is not None else None
        anchor = own.anchor if own is not None and own.anchor is not None \
            and own.anchor.key == key else None
        if anchor is not None and host is self_var and self_id in anchor.ids:
            def cell(f: str, j: int) -> Optional[str]:
                vid = anchor.fields.get(f)
                if vid is None:
                    return None
                if j < 0:
                    return self._opening_ref(vid, f, sheet)
                return self._resolve_var_addr(vid, j, sheet)
            return spell_anchor(field, i, window.grain, start0, cell,
                                spelled_first_day(window))
        vid = None
        if own is not None and own.header_key == key:
            vid = own.fields.get(field)
        if (vid is None or vid == self_id) and anchor is not None:
            vid = anchor.fields.get(field)
        if vid is not None and vid != self_id:
            return self._resolve_var_addr(vid, i, sheet)
        return spell_inline(field, i, window.grain, start0, spelled_first_day(window))

    def _start_cell(self, window: Any, sheet: Optional[str]) -> str:
        """The first period's start as a reference: the window's start
        cell when it is laid out, else the date itself."""
        from ...core.time import _first_day, _parse
        from .header import date_literal
        source = window.start_source
        placed = self.time.placed if self.time is not None else set(self.addresses)
        if source is not None and source.id in placed:
            return self._resolve_var_addr(source.id, 0, sheet)
        first = getattr(window, 'first_day', None)
        return date_literal(first or _first_day(_parse(window.start, window.grain), window.grain))

    def _opening_ref(self, var_id: str, field: str, sheet: Optional[str]) -> Optional[str]:
        """The opening cell of the time row ``var_id`` (field ``field``),
        when the book has an opening column and the row fills it."""
        from .header import OPENING_FIELDS
        addr = self.addresses.get(var_id)
        if field not in OPENING_FIELDS or addr is None or not addr.opening:
            return None
        return self._maybe_qualify(addr.opening, var_id, sheet)

    def render_opening(self, var: Any, sheet: Optional[str]) -> Optional[str | int]:
        """What ``var`` holds in the opening column — formula text, a
        number, or ``None`` for an empty cell: an anchor's end row the day
        before the start (``start - 1``), its period number ``0``, and a
        header mirror of one of them a reference to that cell. Every other row
        leaves the column empty — it has no value before its first
        period."""
        from ...core.time import resolve_default_window, time_field_of
        from .header import OPENING_FIELDS
        var_id = getattr(var, 'id', None)
        addr = self.addresses.get(var_id) if isinstance(var_id, str) else None
        if addr is None or not addr.opening or self.time is None:
            return None
        src = getattr(var, '_mirror_of', None)
        if src is not None:
            return self._opening_ref(src.id, time_field_of(src) or '', sheet)
        field = time_field_of(var)
        own = self.time.sheets.get(sheet)
        if field not in OPENING_FIELDS or own is None or own.anchor is None \
                or var.id not in own.anchor.ids:
            return None
        if field == 'index':
            return 0
        window = resolve_default_window(var)
        if window is None or window.start is None:
            return None
        from ...core.calendar import GREGORIAN
        if GREGORIAN.knows(window.grain) and GREGORIAN.ruler(window.grain).sub_day:
            # "Period 0" of a clock ruler stops where the first period starts.
            return f"{self._start_cell(window, sheet)}"
        return f"{self._start_cell(window, sheet)}-1"

    def render_varref(self, node: VarRef, ctx: RenderCtx) -> str:
        target = self._varref_target(node, ctx)
        if not isinstance(target, Expr):
            return target
        inline = node.var if target is node.var._expr else None
        with self._chain_passes(node, target, ctx), self._entered(inline, ctx):
            return self.walker.render(target, ctx)

    def _varref_target(self, node: VarRef, ctx: RenderCtx) -> Any:
        """What a reference renders as: its text (a cell, a value), or the
        expression of an address-less line, which renders in its place."""
        var_id = self._placed(node.var.id, ctx)
        self_var = ctx.self_var
        if (self_var is not None and node.var is not self_var
                and var_id == getattr(self_var, 'id', None)
                and ctx.self_address is not None
                and self.addresses.get(var_id) is ctx.self_address):
            return self._render_twin(node.var, ctx)
        if ctx.track_role is not None:
            # Inside a track coordinate a TRACKED operand means its
            # SAME-track series (that is what the value was computed
            # from) — never its display-default/blend row.
            coord = self._track_coordinate_ref(node.var, ctx.track_role, ctx)
            if coord is not None:
                return coord
        if var_id in self.addresses:
            dated = self._dated_cell(node, var_id, ctx)
            if dated is not None:
                return dated
            return self._resolve_var_addr(var_id, ctx.period_idx, ctx.current_sheet)
        # An address-less operand may still LIVE in a laid-out row: a
        # loop-built per-month intermediate is the same object as that
        # row's month element. Reference the cell instead of unrolling
        # the expression — except into the cell being written itself,
        # where the reference would be circular and the expression is
        # the honest content.
        home = self.identity_cells.get(id(node.var))
        if home is None and node.var._expr is not None:
            home = self.identity_cells.get(id(node.var._expr))
        if home is not None:
            owner_vid, owner_period = home
            self_var = ctx.self_var
            if not (self_var is not None and owner_vid == self_var.id
                    and owner_period == ctx.period_idx):
                return self._resolve_var_addr(
                    owner_vid, owner_period, ctx.current_sheet
                )
        # ``SUM(x) + a`` with ``x`` a list the book has no cells for:
        # the SUM is one number, written as that number; ``+ a`` stays
        # a reference.
        if isinstance(node.var._expr, FuncCall):
            inlined = self._inline_unlaid_aggregate(node.var._expr, node.var, ctx)
            if inlined is not None:
                return inlined
        if isinstance(node.var._expr, TimeRef):
            return self._render_time(node.var, node.var._expr.field, ctx)
        # Variable not in this emission's layout. Two inlining strategies:
        # - if it has its own AST, recurse into it so we reach real cells;
        # - otherwise (plain input referenced from outside the emission
        #   subtree), inline its value as an Excel literal.
        # Either way, collapsing a LIST into a SCALAR cell positionally
        # is meaningless (SUM over a foreign list would become
        # SUM(first_element)) — mark the render lossy so the cell falls
        # back to the computed value.
        if isinstance(node.var._value, list):
            root = ctx.self_var
            root_is_list = root is not None and (
                getattr(root, 'var_type', None) == 'list'
                or isinstance(getattr(root, '_value', None), list)
            )
            if not root_is_list:
                self._mark_lossy(node, ctx)
        if node is not ctx.own_chain:
            # A roll-forward inlined into another formula: its "previous
            # cell" would be the ENCLOSING row's — a formula that looks
            # right and computes another number. A running total becomes
            # the SUM of its source's cells so far; any other chain ships
            # its value.
            summed = self._render_running_total(node.var, ctx)
            if summed is not None:
                return summed
            if (isinstance(node.var._expr, SelfRef)
                    or getattr(node.var, '_cumsum_source', None) is not None):
                self._mark_lossy(node, ctx)
                return _value_to_excel_literal(
                    _value_at_period(node.var._value, ctx.period_idx))
        if node.var._expr is not None:
            return node.var._expr
        _warn_out_of_scope_ref(node.var, ctx)
        # The number stands in the text (``check_compat`` lists it) — but
        # a Variable outside the model has no row in the workbook, so the
        # number is the formula's for good (no ``inlined_value`` flag).
        literal = _value_to_excel_literal(_value_at_period(node.var._value, ctx.period_idx))
        self.inlined.append((node, literal))
        return literal

    @contextlib.contextmanager
    def _chain_passes(self, node: Any, inner: Any, ctx: RenderCtx):
        """While ``inner`` renders in ``node``'s place — a paren's
        content, an alias's expression — the right to chain on this row's
        previous cell passes to it, and returns after."""
        if node is not ctx.own_chain:
            yield
            return
        ctx.own_chain = inner
        try:
            yield
        finally:
            ctx.own_chain = node

    def _render_running_total(self, var: Any, ctx: RenderCtx) -> Optional[str]:
        """An address-less running total inlined into another formula:
        the SUM of its source's cells from the first period through this
        one. ``None`` where ``var`` is no running total or its source has
        no cells to sum."""
        source = getattr(var, '_cumsum_source', None)
        if source is None:
            return None
        var_id = self._coord_var_id(source, ctx)
        if var_id is None:
            return None
        addr = self.addresses.get(var_id)
        if addr is None or not addr.values or ctx.period_idx >= len(addr.values):
            return None
        cells = [str(a) for a in addr.values[:ctx.period_idx + 1]]
        spans = _run_spans(cells)
        if len(spans) == 1:
            first = self._maybe_qualify(cells[0], var_id, ctx.current_sheet)
            return f"SUM({first}:{cells[-1]})"
        # Its cells are not one run (subtotal columns stand between the
        # months, or a sheet of buckets): one argument per run, never a
        # range that takes in the columns between.
        text = "SUM({})".format(", ".join(
            self._run_text(cells, lo, hi, var_id, ctx.current_sheet) for lo, hi in spans))
        return self._limited(VarRef(var), text, len(spans), [(var_id, cells)], ctx)

    def _render_twin(self, twin: Any, ctx: RenderCtx) -> str:
        """A reference that lands on the cell being written: ``twin``
        shares the row's identity (a projected tree keys its lines like
        its source, ``m.at('quarter')``), so a copy of it would read
        itself. Write what the twin holds instead — its formula when it
        reads only scalars (their cells in this book hold the same
        numbers), else its value: a formula over another grain's rows
        would compute something else."""
        expr = getattr(twin, '_expr', None)
        if expr is not None and not self._reads_series(expr, set()):
            return self.walker.render(expr, ctx)
        self._mark_lossy(VarRef(twin), ctx)
        return _value_to_excel_literal(_value_at_period(twin._value, ctx.period_idx))

    def _reads_series(self, expr: Any, seen: set) -> bool:
        if isinstance(expr, (TimeRef, SelfRef, RollingAggregate, Regrain)):
            return True
        if isinstance(expr, VarRef):
            var = expr.var
            value = getattr(var, '_value', None)
            if isinstance(value, list) or hasattr(value, 'roles'):
                return True
            if var.id in self.addresses or id(var) in seen:
                return False
            seen.add(id(var))
            inner = getattr(var, '_expr', None)
            return inner is not None and self._reads_series(inner, seen)
        return any(self._reads_series(c, seen) for c in _children(expr))

    # ─── The coordinate law ─────────────────────────────────────
    # Inside a track coordinate (``ctx.track_role``) a TRACKED operand
    # means its same-track series — never its display-default/blend
    # row, which computes a different number. Every address lookup in
    # this renderer goes through ``_coord_var_id`` so the rule holds
    # for plain references, lags, ranges, subscripts and rolling
    # windows alike; a coordinate with no laid-out row degrades to the
    # true value (one operand, or the whole cell), never to a formula
    # over the wrong row.

    @staticmethod
    def _is_tracked(var: Any) -> bool:
        value = getattr(var, '_value', None)
        return value is not None and hasattr(value, 'roles')

    def _coord_var_id(self, var: Any, ctx: RenderCtx) -> Optional[str]:
        """The address-book id to read ``var`` through in ``ctx``.

        ``var.id`` outside a coordinate, or for an untracked operand.
        Inside one: the subrow id when that track is laid out; the head
        id when the head IS that track (shows it, or shows the blend
        standing in for it — a follow-only line's live row equals its
        follow track); else None — nothing in the book carries this
        coordinate, the caller must degrade to the value.
        """
        vid = var.id
        role = ctx.track_role
        if role is None or not self._is_tracked(var):
            return self._placed(vid, ctx)
        sub = f"{vid}__track_{role}"
        if sub in self.addresses:
            return sub
        if vid in self.addresses:
            shown = getattr(var, 'display_track_role', lambda: None)()
            if shown == role:
                return vid
            shown_track = getattr(var, 'shown_track', None)
            found = shown_track() if callable(shown_track) else None
            if found is not None and found[0] == role:
                return vid
        return None

    def _coord_value(self, var: Any, role: str, period_idx: int) -> Any:
        value = getattr(var, '_value', None)
        if value is None or not hasattr(value, 'roles') or role not in value:
            return None
        return _value_at_period(value[role], period_idx)

    def _track_coordinate_ref(self, var: Any, role: str,
                              ctx: RenderCtx) -> Optional[str]:
        """A tracked ``var`` read at coordinate ``role``: the cell of its
        laid-out subrow for that track, its head row when the head IS
        that track, else the track's value inlined (the one operand
        degrades, the formula around it stays live). None when ``var``
        carries no tracks, or is an address-less expression to inline —
        the caller resolves it the ordinary way."""
        if not self._is_tracked(var):
            return None
        vid = var.id
        cid = self._coord_var_id(var, ctx) if ctx.track_role == role else None
        if cid is None and ctx.track_role != role:
            # An explicit slice (``x.at(track=r)``) asks for coordinate r
            # regardless of the enclosing one.
            sub = f"{vid}__track_{role}"
            if sub in self.addresses:
                cid = sub
            elif vid in self.addresses:
                shown = getattr(var, 'display_track_role', lambda: None)()
                shown_track = getattr(var, 'shown_track', None)
                found = shown_track() if callable(shown_track) else None
                if shown == role or (found is not None and found[0] == role):
                    cid = vid
        if cid is not None:
            return self._resolve_var_addr(cid, ctx.period_idx, ctx.current_sheet)
        if vid not in self.addresses and getattr(var, '_expr', None) is not None:
            # An address-less EXPRESSION operand (``driver * 2`` inside
            # ``plan=driver * 2``) is inlined by the ordinary path — its
            # own tree keeps rendering inside this coordinate, so the
            # rows it reads resolve to their same-track cells.
            return None
        inlined = self._coord_value(var, role, ctx.period_idx)
        if inlined is None and role not in getattr(var, '_value', {}):
            return None
        return self._note_inlined(Restrict(VarRef(var), role), _value_to_excel_literal(inlined), ctx)

    def render_restrict(self, node: Restrict, ctx: RenderCtx) -> str:
        """``x.at(track='plan')`` — a reference into the coordinate's
        laid-out row.

        The expanded emission tree lays a tracked line out as its
        display-default row (the line's own id) plus one subrow per
        other track, identified ``<id>__track_<role>``. The restrict
        resolves to that subrow's cell; to the head row when the
        restricted role IS the display default (a single-role line, or
        the blend itself); and, when no coordinate row is laid out
        (a blend-only layout), inlines the track's value — the
        one operand degrades, never the whole formula.
        """
        base = node.base
        if isinstance(base, VarRef):
            ref = self._track_coordinate_ref(base.var, node.label, ctx)
            if ref is not None:
                return ref
            value = getattr(base.var, '_value', None)
            if value is not None and hasattr(value, 'roles') and node.label in value:
                return _value_to_excel_literal(
                    _value_at_period(value[node.label], ctx.period_idx)
                )
        return self.walker.render(base, ctx)

    def _inlined_chain_ref(self, node) -> bool:
        """A VarRef to an address-less cumsum result — its inline
        render is a self-referencing chain whose ``prev`` resolves to
        THIS cell's own previous period, which already contains every
        other term of the enclosing formula."""
        return (
            isinstance(node, VarRef)
            and getattr(node.var, '_cumsum_source', None) is not None
            and node.var.id not in self.addresses
        )

    @staticmethod
    def _period_constant(node) -> bool:
        """Renders to the same expression at every period — a scalar
        input or literal; safe to fold into a chain's seed."""
        if isinstance(node, Literal):
            return True
        if isinstance(node, VarRef):
            value = getattr(node.var, '_value', None)
            if isinstance(value, list):
                return False
            # A tracked SERIES varies by period exactly like a list —
            # only an all-scalar tracked value is period-constant.
            if hasattr(value, 'roles'):
                return getattr(value, 'time_length', None) is None
            return True
        return False

    def render_binop(self, node: BinOp, ctx: RenderCtx) -> str:
        # ``scalar + cumsum(x)`` as the WHOLE cell — the linear
        # composition law. The inlined chain's ``prev`` is THIS row's
        # previous cell, which already includes the scalar; re-adding it
        # every period compounds the constant (the double-counted opening
        # balance). Period 0 keeps both terms (seed); later periods are
        # the chain alone. Anywhere else — another operator, a term that
        # moves, the sum inside a larger formula — the chain renders on
        # its own terms: the SUM of its source, or the cell falls back to
        # its computed value.
        folded = self._folded_chain(node, ctx)
        if folded is None:
            return self._render_binop_terms(node, ctx)
        ctx.own_chain = folded
        try:
            if ctx.period_idx == 0:
                return self._render_binop_terms(node, ctx)
            return self.walker.render(folded, ctx)
        finally:
            ctx.own_chain = node

    def _folded_chain(self, node: BinOp, ctx: RenderCtx) -> Optional[Any]:
        """The running total the composition law folds this sum onto —
        ``None`` where the law does not hold."""
        for chain_side, other_side in ((node.left, node.right),
                                       (node.right, node.left)):
            if not self._inlined_chain_ref(chain_side):
                continue
            if (node is ctx.own_chain and node.op == '+'
                    and self._period_constant(other_side)):
                return chain_side
            if self._render_running_total(chain_side.var, ctx) is None:
                self._mark_lossy(chain_side, ctx)
            return None
        return None

    def _render_binop_terms(self, node: BinOp, ctx: RenderCtx) -> str:
        op = node.op
        # // → INT(a/b), % → MOD(a,b). Inside INT/MOD the relevant parent op
        # is ``/`` (precedence 3); the comma in MOD separates and needs no
        # parens-handling on either operand.
        if op == "//":
            left = self._render_operand(node.left, "/", ctx, side="left")
            right = self._render_operand(node.right, "/", ctx, side="right")
            return f"INT({left}/{right})"
        if op == "%":
            left = self.walker.render(node.left, ctx)
            right = self.walker.render(node.right, ctx)
            return f"MOD({left},{right})"

        # ``a + b - c`` is one run down the left side, rendered in a loop:
        # ``sum(rows)`` is an unnamed step per row, the previous step plus
        # that row, and rendered step inside step a few hundred rows
        # outgrew Python's stack — the cell fell back to its number.
        level = next((ops for ops in _CHAIN_LEVELS if op in ops), ())
        chain = [node]
        # Each step of the run down an unnamed intermediate is written in
        # that intermediate's element (``_entered``), its left side read
        # in it too: entered on the way down, left on the way back up.
        entered: List[Any] = []
        try:
            operand: Any = node.left
            while True:
                target = operand
                if isinstance(operand, VarRef) and operand is not ctx.own_chain:
                    target = self._varref_target(operand, ctx)
                if not (isinstance(target, BinOp) and target.op in level
                        and self._folded_chain(target, ctx) is None):
                    break
                chain.append(target)
                step_in = self._entered(operand.var if isinstance(operand, VarRef)
                                        and target is operand.var._expr else None, ctx)
                step_in.__enter__()
                entered.append(step_in)
                operand = target.left
            if target is operand:
                rendered = self.walker.render(operand, ctx)
            elif isinstance(target, Expr):
                last = operand.var if (isinstance(operand, VarRef)
                                       and target is operand.var._expr) else None
                with self._chain_passes(operand, target, ctx), self._entered(last, ctx):
                    rendered = self.walker.render(target, ctx)
            else:
                rendered = target
            parts = [self._wrap_operand(operand, rendered, chain[-1].op, side="left")]
            for k in range(len(chain) - 1, -1, -1):
                while len(entered) > k:
                    entered.pop().__exit__(None, None, None)
                step = chain[k]
                parts.append(_PY_TO_EXCEL_OP.get(step.op, step.op))
                parts.append(self._render_operand(step.right, step.op, ctx, side="right"))
        finally:
            while entered:
                entered.pop().__exit__(None, None, None)
        text = " ".join(parts)
        # A sum of a thousand rows outgrows Excel's formula: as a run
        # of cells it is ``SUM(B2:B1001)``, the way a hand-built book
        # writes it — never any shorter chain, which reads as authored.
        if len(text) >= _MAX_FORMULA_CHARS:
            folded = _fold_sum(parts)
            if folded is not None:
                return folded
        return text

    def _render_operand(
        self, child: Expr, parent_op: str, ctx: RenderCtx, *, side: str
    ) -> str:
        """Render a BinOp operand and wrap in parens iff Excel precedence
        would otherwise re-associate the expression incorrectly.

        Wrapping is decided by the *effective* top-level operator of the
        rendered string — `_effective_op` follows VarRefs into floating
        Variables (whose `_expr` is rendered inline) so we don't miss
        compound subtrees that present as VarRef in the AST.
        """
        return self._wrap_operand(child, self.walker.render(child, ctx), parent_op, side=side)

    def _wrap_operand(self, child: Expr, rendered: str, parent_op: str, *, side: str) -> str:
        effective = self._effective_op(child)
        if effective is None:
            return rendered
        parent_prec = _BINOP_PRECEDENCE.get(parent_op, 0)
        child_prec = _BINOP_PRECEDENCE.get(effective, 0)
        if child_prec < parent_prec:
            return f"({rendered})"
        if (
            child_prec == parent_prec
            and side == "right"
            and parent_op in _NON_COMMUTATIVE_RIGHT
        ):
            return f"({rendered})"
        return rendered

    def _effective_op(self, node: Expr) -> str | None:
        """The operator that would govern this node's rendered string, or
        ``None`` if it renders as an atomic token (cell ref, literal,
        function call, already-parenthesized subexpr).

        Follows VarRefs into floating Variables, since their `_expr` is
        rendered inline at the parent's call site.
        """
        if isinstance(node, BinOp):
            return node.op
        if isinstance(node, Compare):
            # Comparisons (=, >, …) have the lowest precedence, so a
            # comparison used as an arithmetic operand must be parenthesized:
            # ``prev * (hist = 0)``, never ``prev * hist = 0`` (which Excel
            # re-reads as ``(prev * hist) = 0``).
            return node.op
        if isinstance(node, VarRef):
            var = node.var
            if var.id in self.addresses:
                return None  # rendered as a cell address — atomic
            if var._expr is not None:
                return self._effective_op(var._expr)
            return None
        if isinstance(node, MethodCall) and node.method == "copy":
            # A copy renders as what it copies (render_methodcall).
            return self._effective_op(node.base)
        if isinstance(node, ListExpr):
            # Renders as ONE positional item (render_listexpr), but this
            # helper has no period context — report the loosest-binding
            # (minimum-precedence) op among items so the caller adds parens
            # whenever ANY period's item would need them; redundant parens
            # on the other periods are harmless, missing ones re-associate
            # the formula.
            ops = [op for item in node.items
                   if (op := self._effective_op(item)) is not None]
            if not ops:
                return None
            return min(ops, key=lambda o: _BINOP_PRECEDENCE.get(o, 0))
        return None

    def render_subscript(self, node: Subscript, ctx: RenderCtx) -> str:
        if not isinstance(node.base, VarRef):
            # No row holds an element of a bare expression: the whole
            # cell ships its value.
            return self._ship_value(node, ctx)

        var_id = self._coord_var_id(node.base.var, ctx)
        if var_id is None:
            # No row carries this coordinate: element-wise, the whole
            # cell degrades to its computed value.
            return self._ship_value(node, ctx)
        addr_obj = self.addresses.get(var_id)
        if addr_obj is None:
            # An address-less intermediate (``(x * 2.0)[1]``, a cumsum's
            # start): its element's VALUE goes into the text.
            element = self._inline_subscript_value(node.base.var, node.key, ctx.period_idx)
            if element is None:
                return self._ship_value(node, ctx)
            return self._note_inlined(node, element, ctx)

        key = node.key
        if isinstance(key, slice):
            # PERIOD-AWARE element, not a range. A slice in an element-
            # wise formula (``cogs = pl_cogs[:20]`` mapping a 26-period
            # driver onto a 20-period statement) must emit the CURRENT
            # period's cell within the sliced window; emitting the whole
            # range stamped the same ``=Input!B1:U1`` into every cell —
            # #SPILL! chaos in Excel 365, accidental implicit
            # intersection in older Excel. Aggregate args that genuinely want
            # the range (``SUM(x[0:12])``) never reach here —
            # ``_render_range_or_cell`` intercepts them.
            values = addr_obj.values[key]
            if not values:
                return ""
            if len(values) == 1:
                # One-cell window → broadcast: every period reads it.
                return self._maybe_qualify(values[0], var_id, ctx.current_sheet)
            idx = ctx.period_idx
            if idx >= len(values):
                logger.warning(
                    "Slice window on %r has %d cells but period %d is being "
                    "rendered — clamping to the last cell (the owning row is "
                    "longer than the sliced source).",
                    var_id, len(values), idx,
                )
                idx = len(values) - 1
            return self._maybe_qualify(values[idx], var_id, ctx.current_sheet)

        keyed_idx = self._lookup_key_index(node.base.var, key, addr_obj)
        if keyed_idx is not None:
            return self._maybe_qualify(addr_obj.values[keyed_idx], var_id, ctx.current_sheet)

        if isinstance(key, int):
            n = len(addr_obj.values)
            # ``x[-1]`` is the last period, as in Python — on a row whose
            # cells are its own periods.
            idx = key + n if key < 0 and self._is_native_row(node.base.var, addr_obj) else key
            if 0 <= idx < n:
                return self._maybe_qualify(addr_obj.values[idx], var_id, ctx.current_sheet)

        # No cell holds this element (a key the row does not carry):
        # the cell ships its value.
        return self._ship_value(node, ctx)

    def render_methodcall(self, node: MethodCall, ctx: RenderCtx) -> str:
        if node.method == "copy":
            return self.walker.render(node.base, ctx)
        if node.method == "shift":
            return self._render_shift(node, ctx)
        if node.method == "cumsum":
            return self._render_cumsum_chain(node, ctx)

        logger.warning(
            "Method %r on %r has no Excel translation — the cell carries "
            "its computed value.",
            node.method, type(node.base).__name__,
        )
        return self._ship_value(node, ctx)

    def render_funccall(self, node: FuncCall, ctx: RenderCtx) -> str:
        func = node.func.upper()

        # Backend-specific function: FuncCall declared ``render_backends``
        # without including Excel. Can't render natively — inline the
        # owning Variable's value as a literal, the same fallback used
        # by VarRef when its target is outside the current emission.
        if node.render_backends is not None and 'excel' not in node.render_backends:
            return self._fallback_funccall(node, ctx)

        if is_range_taking(func) and not _per_period_call(node, self._lists_series):
            # Reached here, the call is the cell's whole formula (a call
            # inside a larger one is met through its Variable first).
            inlined = self._inline_unlaid_aggregate(node, None, ctx)
            if inlined is not None:
                return inlined
            # ``SUM(a * b)`` over laid-out rows is Excel's SUMPRODUCT.
            rows = self._product_rows(node, ctx)
            if rows is not None:
                return self._sumproduct(node, rows, ctx)
            # A conditional count or sum over rows whose cells are not one
            # run is the sum of one call per run.
            summed = self._per_run(node, func, ctx)
            if summed is not None:
                return summed
            # A range position is the row's range; any other position
            # is this period's cell (a per-period criterion, a key).
            positions = range_positions(func)
            parts = [
                self._render_range_or_cell(a, ctx, func, node)
                if positions is None or i in positions
                else self.walker.render(a, ctx)
                for i, a in enumerate(node.args)
            ]
            text = f"{file_prefix(func)}{excel_name(func)}({', '.join(parts)})"
            if len(text) >= _MAX_FORMULA_CHARS:
                # A row written as its runs can pass what one formula holds.
                reads = [self._row_cells(a, ctx) for i, a in enumerate(node.args)
                         if positions is None or i in positions]
                rows = [read for read in reads if read is not None]
                if any(len(self._runs(cells)) > 1 for _cid, cells in rows):
                    return self._limited(node, text, 0, rows, ctx)
            return text

        if is_range_taking(func):
            # One call per period: a list of lines at a range position is
            # the run of their cells in this period's column.
            positions = range_positions(func)
            parts = [
                self._render_cell_list(a, ctx, func, ctx.period_idx)
                if isinstance(a, ListExpr) and (positions is None or i in positions)
                else self.walker.render(a, ctx)
                for i, a in enumerate(node.args)
            ]
            return f"{file_prefix(func)}{excel_name(func)}({', '.join(parts)})"

        parts = [self.walker.render(a, ctx) for a in node.args]

        if func == "CONCAT":
            # Emit the `&` operator, not the CONCAT() function. CONCAT is an
            # Excel-2016 "future function": in the xlsx it must be stored as
            # ``_xlfn.CONCAT`` or Excel marks it ``@CONCAT`` / ``#NAME?``.
            # ``&`` is universal and is what hand-built models use.
            pieces = [
                f"({p})" if isinstance(arg, (BinOp, Compare)) else p
                for arg, p in zip(node.args, parts)
            ]
            return " & ".join(pieces)

        if is_passthrough(func):
            return f"{file_prefix(func)}{excel_name(func)}({', '.join(parts)})"

        logger.warning(
            "Unknown function %r in formula — emitting as-is. If this is a "
            "valid Excel function, register it in modeleon.functions.registry; "
            "otherwise the cell will be invalid.",
            node.func,
        )
        return f"{node.func}({', '.join(parts)})"

    def _fallback_funccall(self, node: FuncCall, ctx: RenderCtx) -> str:
        """Inline the owning Variable's ``_value`` as an Excel literal
        when a :class:`FuncCall` is declared non-Excel (``render_backends``
        doesn't include ``'excel'``). Warns and emits an empty literal
        if no value is available (backend-exclusive function with no
        Python implementation)."""
        owner = ctx.self_var
        if owner is not None:
            value = _value_at_period(owner._value, ctx.period_idx)
            if value is not None:
                return self._note_inlined(node, _value_to_excel_literal(value), ctx)
        logger.warning(
            "FuncCall %r declared render_backends=%r (excludes 'excel'); "
            "no value available to inline. Cell will be blank.",
            node.func, node.render_backends,
        )
        return '""'

    def render_selfref(self, node: SelfRef, ctx: RenderCtx) -> str:
        # Period 0 is the start value verbatim — no template expansion.
        # Excel cell becomes a direct reference to the start expression
        # (or its inlined literal if the start has no cell address).
        if ctx.period_idx == 0:
            return _strip_outer_parens(self.walker.render(node.start, ctx))

        expansions: Dict[str, str] = {}
        for name, var_expr in node.variables.items():
            text = self.walker.render(var_expr, ctx)
            # ``{prev} * {g}`` with ``g = x + y`` is ``prev * (x + y)``.
            expansions[name] = f"({text})" if self._needs_parens(var_expr) else text

        if ctx.self_address is None or ctx.period_idx - 1 >= len(ctx.self_address.values):
            prev_str = "0"
        else:
            prev_str = ctx.self_address.values[ctx.period_idx - 1]
        expansions["prev"] = prev_str

        # Support both forms the user may write:
        #   ``"{prev} * (1 + {growth})"``  — explicit braces
        #   ``"prev * (1 + growth)"``       — bare identifiers (Python-style)
        # Python evaluation handles bare names natively (AST eval against the
        # variables dict); without the bare form, Excel emission would leave
        # ``prev`` / ``growth`` unsubstituted (``=prev * (1 + growth)``).
        #
        # ONE pass over the template: a substituted reference is never
        # scanned again, so a sheet named ``'Bob''s data'`` or
        # ``'prev year'`` survives a variable named ``s`` or ``prev``.
        # Word boundaries keep ``growth`` out of ``growth_rate``; longest
        # names first so ``g`` doesn't shadow ``growth``. Identifiers may
        # be Unicode — ``{収入}`` is as legal as ``{growth}``.
        names = sorted(expansions, key=len, reverse=True)
        alt = "|".join(re.escape(n) for n in names)
        pattern = re.compile(rf"\{{({alt})\}}|\b({alt})\b")

        def _sub(match: "re.Match[str]") -> str:
            return expansions[match.group(1) or match.group(2)]

        return pattern.sub(_sub, node.template)

    def render_rollingaggregate(self, node: RollingAggregate, ctx: RenderCtx) -> str:
        if ctx.period_idx < node.window - 1:
            return _value_to_excel_literal(node.fill)

        var_id = self._coord_var_id(node.source, ctx)
        if var_id is None:
            self._mark_lossy(node, ctx)
            return _value_to_excel_literal(node.fill)
        addr_obj = self.addresses.get(var_id)
        if addr_obj is None or not addr_obj.values:
            # No cells to take the window over: the cell carries its value.
            self._mark_lossy(node, ctx)
            return _value_to_excel_literal(node.fill)

        start_idx = ctx.period_idx - node.window + 1
        end_idx = min(ctx.period_idx, len(addr_obj.values) - 1)
        # The window's own cells, never one range from the first to the
        # last: a declared quarter total stands between March and April,
        # and SUM over Feb..Apr would add it in.
        cells = [str(a) for a in addr_obj.values[start_idx:end_idx + 1]]
        span = self._cells_as_range(cells, var_id, ctx.current_sheet, ctx, node.func, node)
        text = f"{node.func}({span})"
        if len(text) >= _MAX_FORMULA_CHARS:
            return self._limited(node, text, 0, [(var_id, cells)], ctx)
        return text

    @staticmethod
    def _is_native_row(source: Any, addr_obj: Any) -> bool:
        """Do these cells hold ``source``'s own periods, one each?"""
        value = getattr(source, '_value', None)
        if hasattr(value, 'roles'):
            length = getattr(value, 'time_length', None)
        elif isinstance(value, list):
            length = len(value)
        else:
            return True
        return length is None or length == len(addr_obj.values)

    def render_regrain(self, node: Regrain, ctx: RenderCtx) -> str:
        i = ctx.period_idx
        # An un-entered period is a blank cell, zero in the fold, as the
        # engine's re-grain counts a hole - but in a line of dates or
        # words the engine's bucket stays blank, and so does the cell.
        fill = node.fill_values[i] if i < len(node.fill_values) else 0.0
        if fill is None:
            return '""'
        var_id = self._coord_var_id(node.source, ctx)
        if var_id is None:
            return self._note_inlined(node, _value_to_excel_literal(fill), ctx)
        addr_obj = self.addresses.get(var_id)
        if (addr_obj is None or not addr_obj.values or i >= len(node.buckets)):
            return self._note_inlined(node, _value_to_excel_literal(fill), ctx)
        if addr_obj is ctx.self_address or not self._is_native_row(node.source, addr_obj):
            # The cells under the source's id are not its native periods:
            # a projected book keys its rows like the source, so they are
            # this very row or another grain's. The bucket's value is the
            # truth; a range over those cells would read the wrong periods.
            return self._note_inlined(node, _value_to_excel_literal(fill), ctx)
        lo, hi = node.buckets[i]
        n = len(addr_obj.values)
        if lo >= n or hi - 1 >= n:
            return self._note_inlined(node, _value_to_excel_literal(fill), ctx)
        from .totals import bucket_formula
        # The bucket's own cells, never one range from the first to the
        # last: subtotal columns may stand between them. The engine's
        # fold, spelled cell for cell (a mean weighted by the days).
        weights = (self._bucket_weights(node.source, lo, hi)
                   if node.recipe == 'mean' else None)
        series = getattr(node.source, '_value', None)
        if ctx.track_role is not None and hasattr(series, 'roles'):
            series = series[ctx.track_role] if ctx.track_role in series.roles else None
        body = bucket_formula(
            node.recipe, list(addr_obj.values[lo:hi]), weights,
            qualify=lambda ref: self._maybe_qualify(ref, var_id, ctx.current_sheet),
            blanks=not isinstance(series, list) or any(v is None for v in series[lo:hi]),
        )
        if body is None:
            return _value_to_excel_literal(fill)
        # A mean over the days ``(a*31+b*28)/59`` or a compounded rate
        # ``EXP(...)-1`` is an expression, not a token: as an operand it
        # keeps its brackets, or ``K / mean`` divided by the days again.
        return f"({body})" if _has_top_level_operator(body) else body

    @staticmethod
    def _bucket_weights(source: Any, lo: int, hi: int) -> Optional[List[int]]:
        """The day counts of ``source``'s native periods ``lo``..``hi`` —
        the weights its mean folds by (:func:`period_weights`)."""
        loc = getattr(source, 'time', None)
        if loc is None or loc.start is None or loc.grain is None:
            return None
        from ...core.time import period_weights
        try:
            return period_weights(loc.grain, loc.start, hi, loc.first_day)[lo:hi]
        except (ValueError, TypeError):
            return None

    # ─── Excel-specific helpers ─────────────────────────────────

    def _needs_parens(self, inner: Expr) -> bool:
        if isinstance(inner, (BinOp, UnaryOp, Compare)):
            return True
        if isinstance(inner, MethodCall) and inner.method == "copy":
            # A copy renders as what it copies (a constant step carried
            # into a coarser grain): bracketed as that is.
            return self._needs_parens(inner.base)
        if isinstance(inner, VarRef):
            var = inner.var
            if var.id in self.addresses:
                return False
            if var._expr is not None:
                return self._needs_parens(var._expr)
            return False
        if isinstance(inner, ListExpr):
            # Renders as one positional item; parenthesize if ANY period's
            # item would need it (no period context here — see
            # _effective_op).
            return any(self._needs_parens(item) for item in inner.items)
        return False

    def _render_cumsum_chain(self, node: MethodCall, ctx: RenderCtx) -> str:
        """``MethodCall(VarRef(x), 'cumsum')`` — the rank-lifted spelling.

        ``mo.cumsum`` over a TRACKED Variable keeps its formula symbolic
        (values rank-lift per track) instead of delegating to
        ``recurrence``, so the Excel renderer sees a method call. Emit
        the same running chain the recurrence path produces::

            period 0:  =<input cell>
            period t:  =<own cell at t-1> + <input cell at t>

        ``ctx.self_address.values`` is period-ordered for the exact row
        being written — under a tracks layout that's the per-track cell
        list, so ``values[t-1]`` lands on the previous period of the
        SAME track regardless of the layout's column stride.

        Without an own address (the cumsum is inlined inside a larger
        formula) a chain has no cell for ``prev`` to point at — mark
        the render lossy so the cell ships its computed value, never a
        formula Excel can't evaluate.
        """
        base = self.walker.render(node.base, ctx)
        if ctx.period_idx == 0:
            return base
        if (
            ctx.self_address is not None
            and ctx.period_idx - 1 < len(ctx.self_address.values)
        ):
            prev = ctx.self_address.values[ctx.period_idx - 1]
            return f"{prev} + {base}"
        return self._ship_value(node, ctx)

    def _render_shift(self, node: MethodCall, ctx: RenderCtx) -> str:
        if not isinstance(node.base, VarRef):
            return self._ship_value(node, ctx)

        periods = node.args[0] if node.args else 1
        try:
            periods_int = int(periods)
        except (TypeError, ValueError):
            periods_int = 1

        fill_value = node.kwargs.get("fill_value", 0)
        target_period = ctx.period_idx - periods_int

        def _fill() -> str:
            # A Variable-backed fill (``mo.lag(x, fill=opening)``)
            # arrives as an Expr — render it as a reference to the
            # fill cell so the dependency stays live in the workbook.
            # An unnamed compound fill (``fill=a + b - c``) is inlined,
            # and the lag reads as ONE operand wherever it sits:
            # ``x - lag(y, fill=a + b - c)`` is ``B2 - (B1 + B2 - B3)``.
            if isinstance(fill_value, Expr):
                text = self.walker.render(fill_value, ctx)
                return f"({text})" if self._needs_parens(fill_value) else text
            return _value_to_excel_literal(fill_value)

        if target_period < 0:
            return _fill()
        base_expr = getattr(node.base.var, '_expr', None)
        if isinstance(base_expr, TimeRef) and node.base.var.id not in self.addresses:
            base_values = getattr(node.base.var, '_value', None)
            if isinstance(base_values, list) and target_period >= len(base_values):
                return _fill()
            # ``mo.lag(mo.time.end)`` — the previous period's time.
            here = ctx.period_idx
            ctx.period_idx = target_period
            try:
                return self._render_time(node.base.var, base_expr.field, ctx)
            finally:
                ctx.period_idx = here
        cid = self._coord_var_id(node.base.var, ctx)
        if cid is None:
            # The coordinate has no laid-out row: the lagged VALUE of
            # that track is the honest content (one operand degrades).
            shifted = self._coord_value(node.base.var, ctx.track_role or "", target_period)
            if shifted is None:
                return _fill()
            return self._note_inlined(node, _value_to_excel_literal(shifted), ctx)
        addr_obj = self.addresses.get(cid)
        if addr_obj is None:
            return self._shift_unlaid(node.base.var, target_period, _fill, ctx)
        if target_period >= len(addr_obj.values):
            return _fill()

        return self._maybe_qualify(addr_obj.values[target_period], cid, ctx.current_sheet)

    def _shift_unlaid(self, base: Any, target_period: int, fill: Callable[[], str],
                      ctx: RenderCtx) -> str:
        """``mo.lag(a.x + a.y)`` — the base has no row of its own, so the
        lagged cell is the base's formula written for ``target_period``.
        A base whose formula leans on the cell being written (a
        recurrence's ``{prev}``, a running sum, a value fallback) cannot
        be moved to another period: the whole cell ships its value."""
        values = getattr(base, '_value', None)
        length = (getattr(values, 'time_length', None) if hasattr(values, 'roles')
                  else len(values) if isinstance(values, list) else None)
        if length is not None and target_period >= length:
            return fill()
        if self._reads_own_cell(base, set()):
            self._mark_lossy(VarRef(base), ctx)
            return fill()
        here = ctx.period_idx
        ctx.period_idx = target_period
        try:
            shifted = self.walker.render(VarRef(base), ctx)
        finally:
            ctx.period_idx = here
        # The lag reads as ONE operand wherever it sits:
        # ``mo.lag(x + y) * 2`` is ``(B1 + B2) * 2``.
        return f"({shifted})" if self._needs_parens(VarRef(base)) else shifted

    def _reads_own_cell(self, var: Any, seen: set) -> bool:
        if id(var) in seen or var.id in self.addresses:
            return False
        seen.add(id(var))
        if id(var) in self.identity_cells:
            return True
        expr = getattr(var, '_expr', None)
        return expr is not None and self._expr_reads_own_cell(expr, seen)

    def _expr_reads_own_cell(self, expr: Any, seen: set) -> bool:
        if isinstance(expr, SelfRef):
            return True
        if isinstance(expr, MethodCall) and expr.method == 'cumsum':
            return True
        if isinstance(expr, FuncCall) and expr.render_backends is not None \
                and 'excel' not in expr.render_backends:
            return True
        if isinstance(expr, VarRef):
            return self._reads_own_cell(expr.var, seen)
        if isinstance(expr, (RollingAggregate, Regrain)):
            return self._reads_own_cell(expr.source, seen)
        return any(self._expr_reads_own_cell(c, seen) for c in _children(expr))

    def _render_range_or_cell(
        self, arg: Expr, ctx: RenderCtx, func: Optional[str], call: FuncCall,
    ) -> str:
        if isinstance(arg, ListExpr):
            block = self._render_block(arg, ctx)
            return self._render_cell_list(arg, ctx, func) if block is None else block
        row = self._row_cells(arg, ctx)
        if row is not None:
            # A row's cells, or a window's — which can step over a totals
            # column too (``x[0:4]`` = Jan–Apr crosses the Q1 total).
            cid, cells = row
            return self._cells_as_range(cells, cid, ctx.current_sheet, ctx, func, call)
        window = self._window(arg)
        if (isinstance(arg, VarRef) and self._coord_var_id(arg.var, ctx) is None) or (
                window is not None and self._coord_var_id(window[0], ctx) is None):
            # A range over a coordinate nothing in the book carries: a
            # per-period literal is not a range, so the whole aggregate
            # degrades to its computed value.
            self._mark_lossy(arg, ctx)
        return self.walker.render(arg, ctx)

    def _window(self, arg: Expr) -> Optional[tuple[Any, slice]]:
        """A SLICED row inside an aggregate wants the sliced RANGE:
        ``SUM(x[0:12])`` → ``SUM(B1:M1)``. (Element-wise slice rendering
        lives in ``render_subscript`` and is period-aware; this is the one
        context where the whole window is the point.) The slice arrives
        either as a bare Subscript node or — the common DSL shape — as a
        VarRef to the intermediate Variable ``x[0:12]`` produced (its
        ``_expr`` is the Subscript). The row sliced, and the slice."""
        sub = arg if isinstance(arg, Subscript) else None
        if (
            sub is None
            and isinstance(arg, VarRef)
            and arg.var.id not in self.addresses
            and isinstance(getattr(arg.var, "_expr", None), Subscript)
        ):
            sub = arg.var._expr
        if (
            sub is not None
            and isinstance(sub.base, VarRef)
            and isinstance(sub.key, slice)
        ):
            return sub.base.var, sub.key
        return None

    def _range_addr(self, var_id: str) -> Optional[VariableAddresses]:
        """The cells of a line read as a range: its own periods (see
        ``periods``)."""
        return self.periods.get(var_id) or self.addresses.get(var_id)

    def _row_cells(self, arg: Expr, ctx: RenderCtx) -> Optional[tuple[str, List[str]]]:
        """The id and the cells of a range argument written as a range of
        its row's cells (``_render_range_or_cell``) — a laid-out row, or a
        window of one; None for anything else."""
        if isinstance(arg, VarRef):
            cid = self._coord_var_id(arg.var, ctx)
            if cid is None:
                return None
            addr_obj = self._range_addr(cid)
            if addr_obj is not None and len(addr_obj.values) > 1:
                return cid, addr_obj.values
        window = self._window(arg)
        if window is not None:
            base, key = window
            cid = self._coord_var_id(base, ctx)
            addr_obj = self._range_addr(cid) if cid is not None else None
            if cid is not None and addr_obj is not None and len(addr_obj.values[key]) > 1:
                return cid, addr_obj.values[key]
        return None

    def _runs(self, cells: List[str]) -> List[tuple[int, int]]:
        """The runs of these cells (:func:`_run_spans`), split once per
        render: nothing a render reads moves while it lasts."""
        key = tuple(cells)
        spans = self._spans.get(key)
        if spans is None:
            spans = self._spans[key] = _run_spans(cells)
        return spans

    def _split(self, reads: Iterable[tuple[str, List[str]]],
               ) -> tuple[tuple[Optional[str], Optional[str], tuple[str, ...]], ...]:
        """Of the rows read (id and cells), those whose cells are not one
        run — with their sheet, as :class:`NotOneRun` names them."""
        return tuple((self.var_to_sheet.get(cid), cid, tuple(cells))
                     for cid, cells in reads if len(self._runs(cells)) > 1)

    def _run_text(self, cells: List[str], lo: int, hi: int, var_id: str,
                  current_sheet: Optional[str]) -> str:
        """Cells ``lo``..``hi`` of a line, one run, as a range — its sheet
        named once, in front — or the one cell."""
        first = self._maybe_qualify(cells[lo], var_id, current_sheet)
        return first if hi - lo == 1 else f"{first}:{cells[hi - 1]}"

    def _per_run(self, call: FuncCall, func: str, ctx: RenderCtx) -> Optional[str]:
        """``COUNTIF`` / ``SUMIF`` and their ``S`` forms over rows whose
        cells are not one run: the sum of one call per run, every range cut
        to the same periods, each criteria written once and read by every
        call; ``AVERAGEIF(S)`` the ratio of those sums to the count of the
        numbers they add. None where the ranges are one run, or a range is
        not a row (a list of lines is one range) or not as long as the
        others."""
        at = _criteria_ranges(func, len(call.args))
        if at is None:
            return None
        rows: Dict[int, tuple[str, List[str]]] = {}
        for i in at:
            read = self._row_cells(call.args[i], ctx)
            if read is None:
                return None
            rows[i] = read
        if len({len(cells) for _cid, cells in rows.values()}) != 1:
            return None
        spans = _united_spans([cells for _cid, cells in rows.values()], self._runs)
        if len(spans) == 1:
            return None
        if len(spans) > _MAX_ARGS:
            # Counted before a call per run is written: past what one
            # call takes, the text is never read.
            return self._ship_value(
                call, ctx, NotOneRun("arguments", self._split(rows.values()), len(spans)))
        # Each criteria once: its text stands in every call.
        written = {i: self.walker.render(arg, ctx)
                   for i, arg in enumerate(call.args) if i not in rows}
        sheet = ctx.current_sheet

        def one(name: str, lo: int, hi: int, order: List[Any]) -> str:
            # ``order``: the call's arguments by position, a text as it is.
            parts = [i if isinstance(i, str)
                     else self._run_text(rows[i][1], lo, hi, rows[i][0], sheet) if i in rows
                     else written[i] for i in order]
            return f"{file_prefix(name)}{excel_name(name)}({', '.join(parts)})"

        every = list(range(len(call.args)))
        if func in ("AVERAGEIF", "AVERAGEIFS"):
            # The mean of the numbers matched: their sum over their count —
            # a blank, a text or a boolean in the averaged range is not
            # counted, as the mean skips it. The averaged range leads the
            # count, which keeps one criteria pair clear of LibreOffice's
            # COUNTIFS misread of a blank cell at either end of a range;
            # with two or more pairs a blank at a run's end of the averaged
            # row is still misread there (Excel reads it right).
            single = func == "AVERAGEIF"
            averaged = 2 if single and len(every) > 2 else 0
            counted = [averaged, _NUMBERS, *(every[:2] if single else every[1:])]
            total = "SUM({})".format(", ".join(
                one("SUMIF" if single else "SUMIFS", lo, hi, every) for lo, hi in spans))
            count = "SUM({})".format(", ".join(
                one("COUNTIFS", lo, hi, counted) for lo, hi in spans))
            text = f"({total} / {count})"
        else:
            text = "SUM({})".format(", ".join(one(func, lo, hi, every) for lo, hi in spans))
        return self._limited(call, text, len(spans), rows.values(), ctx)

    def _render_cell_list(
        self, node: ListExpr, ctx: RenderCtx, func: Optional[str],
        period: Optional[int] = None,
    ) -> str:
        """A list of scalar lines as one range argument: their cells, one
        under another in a column (a field down the records of a table)
        or side by side on a row, are the range ``Deals!C3:C5``. Cells
        that form no run cannot be one range — a comma list is read as
        further arguments by every function but the variadic ones — so
        the cell carries its computed value. Given a ``period``, a line
        holding a series stands for its cell of that period: a field of
        yearly lines down the records is the run in that year's column."""
        # A period's cells are those of the element being written: read by
        # date or by position as Python read them (``_frame``).
        frame = self._frame(ctx) if period is not None else None
        key = (id(node), ctx.current_sheet, ctx.track_role, period,
               func in self._LIST_SAFE_FUNCS,
               None if frame is None else (frame.offset, frame.window,
                                           tuple(getattr(frame.var, '_keys', None) or ())))
        hit = self._ranges.get(key)
        if hit is not None and hit[0] is node:
            return self._ship_value(node, ctx, hit[2]) if hit[1] is None else hit[1]
        text, cause = self._cell_list_text(node, ctx, func, period)
        if not self._mirrored(node, ctx):
            self._ranges[key] = (node, text, cause)
        return self._ship_value(node, ctx, cause) if text is None else text

    # Does a list of lines hold a series — read once per render.
    def _lists_series(self, node: ListExpr) -> bool:
        hit = self._listed.get(id(node))
        if hit is None or hit[0] is not node:
            hit = self._listed[id(node)] = (node, _lists_series(node))
        return hit[1]

    # Is a line of ``node`` mirrored into this sheet's header? Its cell then
    # depends on the formula being rendered (the mirror reads the row
    # itself), so the text is read for each formula.
    def _mirrored(self, node: ListExpr, ctx: RenderCtx) -> bool:
        book = self.time
        own = book.sheets.get(ctx.current_sheet) if book is not None else None
        if own is None or not own.alias:
            return False
        return any(isinstance(item, VarRef) and item.var.id in own.alias
                   for item in node.items)

    # The text of ``_render_cell_list``, or None where the cell carries its
    # computed value — then with the cause, where its cells are not one run.
    def _cell_list_text(
        self, node: ListExpr, ctx: RenderCtx, func: Optional[str],
        period: Optional[int],
    ) -> tuple[Optional[str], Optional[NotOneRun]]:
        cells: List[tuple[str, str]] = []
        for item in node.items:
            found = self._cell_of(item, ctx, period)
            if found is None:
                return None, None
            cells.append(found)
        refs = [cell for cell, _vid in cells]
        sheets = {self.var_to_sheet.get(vid) for _cell, vid in cells}
        first, vid = cells[0]
        if len(sheets) == 1 and _contiguous_run(refs):
            start = self._maybe_qualify(first, vid, ctx.current_sheet)
            return (start if len(refs) == 1 else f"{start}:{refs[-1]}"), None
        if func in self._LIST_SAFE_FUNCS:
            return ", ".join(
                self._maybe_qualify(cell, vid, ctx.current_sheet) for cell, vid in cells
            ), None
        sheet = sheets.pop() if len(sheets) == 1 else None
        return None, NotOneRun(None, ((sheet, None, tuple(refs)),))

    def _render_block(self, node: ListExpr, ctx: RenderCtx) -> Optional[str]:
        """A list of rows as one rectangular range argument — the records
        of a table, each a list of its cells behind a literal label, or
        laid-out rows one under another — ``Deals!A3:C5``, ``Inputs!B2:D3``.
        A record's label is the cell before its first field, in the
        sheet's first column, where the records layout writes it; with
        anything between the two (a unit column) no cell can be named for
        it. None when ``node`` is a list of cells, not of rows; when the
        rows form no rectangle on one sheet the cell carries its computed
        value."""
        from openpyxl.utils.cell import coordinate_to_tuple

        rows: List[List[str]] = []
        owners: List[str] = []
        series: List[tuple[str, List[str]]] = []
        for item in node.items:
            if isinstance(item, ListExpr):
                cells = item.items
                labelled = bool(cells) and isinstance(cells[0], Literal) and isinstance(
                    cells[0].value, str)
                refs: List[str] = []
                for sub in cells[1 if labelled else 0:]:
                    found = self._cell_of(sub, ctx)
                    if found is None:
                        return self._ship_value(node, ctx)
                    refs.append(found[0])
                    owners.append(found[1])
                if not refs:
                    return self._ship_value(node, ctx)
                if labelled:
                    row_no, col = coordinate_to_tuple(refs[0])
                    if col != 2:
                        return self._ship_value(node, ctx)
                    refs.insert(0, f"A{row_no}")
                rows.append(refs)
            elif isinstance(item, VarRef) and (
                isinstance(item.var._value, list) or hasattr(item.var._value, "roles")
            ):
                vid = self._coord_var_id(item.var, ctx)
                addr = self._range_addr(vid) if vid is not None else None
                if vid is None or addr is None:
                    return self._ship_value(node, ctx)
                rows.append(list(addr.values))
                owners.append(vid)
                series.append((vid, rows[-1]))
            else:
                return None
        if not rows:
            return None
        if len({self.var_to_sheet.get(vid) for vid in owners}) != 1:
            return self._ship_value(node, ctx)
        grid = [[coordinate_to_tuple(ref) for ref in refs] for refs in rows]
        cols = [c for _r, c in grid[0]]
        tops = [row[0][0] for row in grid]
        if (
            any([c for _r, c in row] != cols for row in grid)
            or any(len({r for r, _c in row}) != 1 for row in grid)
            or any(b - a != 1 for a, b in zip(tops, tops[1:]))
        ):
            return self._ship_value(node, ctx)
        if any(b - a != 1 for a, b in zip(cols, cols[1:])):
            # Rows one under another, in the same columns: only their
            # runs keep the block from being one range.
            split = self._split(series)
            return self._ship_value(node, ctx, NotOneRun(None, split) if split else None)
        first = self._maybe_qualify(rows[0][0], owners[0], ctx.current_sheet)
        last = rows[-1][-1]
        return first if rows[0][0] == last else f"{first}:{last}"

    def _cell_of(self, item: Expr, ctx: RenderCtx,
                 period: Optional[int] = None) -> Optional[tuple[str, str]]:
        """The one cell ``item`` stands for — a scalar line's, one element
        of a laid-out row (``flow[0]``), or, given a ``period``, a row's
        cell of that period, as Python read it (``_read_index``: of the
        same date where lines on other windows were aligned) — with the id
        it is addressed through; None when no cell holds it."""
        if isinstance(item, VarRef):
            var_id = self._coord_var_id(item.var, ctx)
            addr_obj = self.addresses.get(var_id) if var_id is not None else None
            if addr_obj is not None and var_id is not None:
                cells = addr_obj.values
                read = (self._read_index(item.var, var_id, addr_obj, ctx)
                        if period is not None and isinstance(item.var._value, list) else None)
                if read is not None and read[1]:
                    return (cells[read[0]], var_id) if 0 <= read[0] < len(cells) else None
                if len(cells) == 1:
                    return cells[0], var_id
                if read is not None and 0 <= read[0] < len(cells):
                    return cells[read[0]], var_id
                return None
            inner = getattr(item.var, "_expr", None)
            if isinstance(inner, Subscript):
                item = inner
        if (
            isinstance(item, Subscript)
            and isinstance(item.base, VarRef)
            and isinstance(item.key, int)
            and not isinstance(item.key, bool)
        ):
            var_id = self._coord_var_id(item.base.var, ctx)
            addr_obj = self.addresses.get(var_id) if var_id is not None else None
            if addr_obj is not None and var_id is not None:
                n = len(addr_obj.values)
                key = item.key
                idx = key + n if key < 0 and self._is_native_row(item.base.var, addr_obj) else key
                if 0 <= idx < n:
                    return addr_obj.values[idx], var_id
        return None

    def _unlaid_list_arg(self, call: FuncCall, ctx: RenderCtx) -> Optional[Expr]:
        """The first list argument of ``call`` that has no cells in this
        workbook — a Variable never attached to the model, an
        intermediate such as ``x * 2`` or ``x[0:2]`` of one, a literal
        ``[10, 20, 30]`` — or None when every list can be written as a
        range. Scalar arguments never count: one value is one cell or
        one literal either way, and nor does a list at a position the
        registry writes as a cell."""
        positions = range_positions(call.func.upper())
        for i, arg in enumerate(call.args):
            if positions is not None and i not in positions:
                continue
            if isinstance(arg, Literal):
                if isinstance(arg.value, list):
                    return arg
                continue
            if not isinstance(arg, VarRef) or not isinstance(arg.var._value, list):
                continue
            cid = self._coord_var_id(arg.var, ctx)
            if cid is None or cid in self.addresses:
                continue
            sub = arg.var._expr
            if (
                isinstance(sub, Subscript)
                and isinstance(sub.key, slice)
                and isinstance(sub.base, VarRef)
            ):
                base_id = self._coord_var_id(sub.base.var, ctx)
                if base_id is None or base_id in self.addresses:
                    continue            # a window of a laid-out row
            if self._product_rows(call, ctx) is not None:
                continue                # SUM of a product of laid-out rows
            return arg
        return None

    def _product_rows(self, call: FuncCall,
                      ctx: RenderCtx) -> Optional[List[tuple[str, List[str]]]]:
        """``SUM(a * b * …)`` over laid-out rows of one period range: each
        row's id and cells, in order — Excel's ``SUMPRODUCT`` over them.
        None for any other call: a scalar factor, a row the book has no
        cells for, a product aligned across windows."""
        if call.func.upper() != "SUM" or len(call.args) != 1:
            return None
        arg = call.args[0]
        if (
            not isinstance(arg, VarRef)
            or not isinstance(arg.var._value, list)
            or getattr(arg.var, "_start", None) is not None
            or self._coord_var_id(arg.var, ctx) in self.addresses
        ):
            return None
        factors: List[Any] = []
        if not self._product_factors(arg.var._expr, factors, ctx):
            return None
        rows: List[tuple[str, List[str]]] = []
        widths: set[int] = set()
        for var in factors:
            cid = self._coord_var_id(var, ctx)
            addr = self._range_addr(cid) if cid is not None else None
            if (
                cid is None or addr is None or len(addr.values) < 2
                or not self._is_native_row(var, addr)
            ):
                return None
            widths.add(len(addr.values))
            rows.append((cid, list(addr.values)))
        if len(widths) != 1 or len(arg.var._value) not in widths:
            return None
        return rows

    def _sumproduct(self, call: FuncCall, rows: List[tuple[str, List[str]]],
                    ctx: RenderCtx) -> str:
        """``SUMPRODUCT`` over the rows of ``SUM(a * b)``; over rows whose
        cells are not one run, the sum of one ``SUMPRODUCT`` per run, every
        row cut to the same periods."""
        spans = _united_spans([cells for _cid, cells in rows], self._runs)
        sheet = ctx.current_sheet

        def one(lo: int, hi: int) -> str:
            return "SUMPRODUCT({})".format(", ".join(
                self._run_text(cells, lo, hi, cid, sheet) for cid, cells in rows))

        if len(spans) == 1:
            return one(*spans[0])
        if len(spans) > _MAX_ARGS:
            return self._ship_value(
                call, ctx, NotOneRun("arguments", self._split(rows), len(spans)))
        text = "SUM({})".format(", ".join(one(lo, hi) for lo, hi in spans))
        return self._limited(call, text, len(spans), rows, ctx)

    def _product_factors(self, expr: Any, out: List[Any], ctx: RenderCtx) -> bool:
        """Collect the rows of a ``*`` chain into ``out``; False where a
        term is anything but a laid-out row or another such product."""
        if isinstance(expr, Paren):
            return self._product_factors(expr.inner, out, ctx)
        if isinstance(expr, BinOp) and expr.op == "*":
            return (self._product_factors(expr.left, out, ctx)
                    and self._product_factors(expr.right, out, ctx))
        if isinstance(expr, VarRef):
            var = expr.var
            if self._coord_var_id(var, ctx) in self.addresses:
                out.append(var)
                return True
            inner = getattr(var, "_expr", None)
            if (isinstance(var._value, list) and inner is not None
                    and getattr(var, "_start", None) is None):
                return self._product_factors(inner, out, ctx)
        return False

    def _inline_unlaid_aggregate(
        self, call: FuncCall, owner: Any, ctx: RenderCtx,
    ) -> Optional[str]:
        """What to write in place of ``SUM(x)`` / ``NPV(r, x)`` / … when a
        list it reduces has no cells in this workbook; None when every
        list has cells.

        Such a list can only be written one period's element at a time,
        and ``SUM`` over one element is not ``SUM`` over the list: the
        workbook would show 11, 22, 33 where the model says 61, 62, 63.
        The call's computed value is one number, the same in every
        period. Inside a larger formula (``owner`` is the Variable the
        call belongs to) that number takes the call's place and the rest
        of the formula stays live. When the call is the cell's whole
        formula (``owner`` None) the cell carries its computed value.
        """
        if not is_range_taking(call.func.upper()) or _per_period_call(call, self._lists_series):
            return None
        arg = self._unlaid_list_arg(call, ctx)
        if arg is None:
            return None
        _warn_unlaid_aggregate(call.func.upper(), arg, ctx)
        value = getattr(owner, "_value", None)
        if (value is None or isinstance(value, list)
                or hasattr(value, "roles")):
            # No single number stands for the call: the whole cell
            # falls back to its computed value.
            return self._ship_value(call, ctx)
        # A number baked into formula text: true now, stale once the
        # list changes — callers that cache formula text must not reuse it.
        return self._note_inlined(call, _value_to_excel_literal(value), ctx)

    def _inline_subscript_value(self, base_var: Any, key: Any,
                                period_idx: int) -> Optional[str]:
        """The element ``base_var[key]`` as a literal — of a window, this
        period's element — or None when the value holds no such element."""
        value = base_var._value
        keys = getattr(base_var, "_keys", None)

        if isinstance(key, slice):
            # Element-wise, as a window of a laid-out row: this period's
            # element; a one-element window broadcasts, a longer row
            # clamps to the last.
            window = value[key] if isinstance(value, list) else []
            if not window:
                return None
            return _value_to_excel_literal(window[min(period_idx, len(window) - 1)])

        if isinstance(key, str) and keys and key in keys:
            return _value_to_excel_literal(value[keys.index(key)])
        if (
            isinstance(key, int)
            and keys
            and key in keys
            and key not in range(len(value) if isinstance(value, list) else 1)
        ):
            return _value_to_excel_literal(value[keys.index(key)])
        if isinstance(key, int) and isinstance(value, list) and -len(value) <= key < len(value):
            return _value_to_excel_literal(value[key])
        return None

    def _lookup_key_index(self, base_var: Any, key: Any, addr_obj: VariableAddresses) -> Optional[int]:
        keys = getattr(base_var, "_keys", None)
        if not keys:
            return None
        if isinstance(key, str):
            if key in keys:
                return keys.index(key)
            return None
        if isinstance(key, int):
            if key in keys and key not in range(len(addr_obj.values)):
                return keys.index(key)
        return None

    def _dated_cell(self, node: VarRef, var_id: str, ctx: RenderCtx) -> Optional[str]:
        """The cell of ``node``'s line this formula reads, where it is not
        the one of the formula's own column: the cell Python read for the
        value being written (``_frame``). By DATE where two windows were
        aligned — an operator or a one-call-per-period function over a
        line on a window of its own inside the sheet's (a tranche from
        March in a model from January: its cell of the same date, as
        ``_onto_window`` reads it) — a period the line has no column for
        continued by its declared ``extend=`` rule: ``0`` for a flow, its
        first or last cell for a rate held, the value where no rule
        continues it. By KEY where lines along the same keys in another
        order were combined; by position otherwise — the column's cell,
        so a formula over one other window, which Python read by
        position, reads it by position too. ``None`` where that is the
        cell of the formula's own column."""
        if ctx.frame is None and _by_position(ctx.self_var):
            return None                 # the column's own cell, as before
        if not isinstance(getattr(node.var, '_value', None), list):
            return None
        addr_obj = self.addresses.get(var_id)
        if addr_obj is None:
            return None
        read = self._read_index(node.var, var_id, addr_obj, ctx)
        if read is None:
            return None
        j, dated = read
        cells = addr_obj.values
        if 0 <= j < len(cells):
            if j == ctx.period_idx:
                return None
            return self._maybe_qualify(cells[j], var_id, ctx.current_sheet)
        if not dated:
            return None
        rule = getattr(node.var, '_extend', None)
        if rule == 'zero':
            return "0"
        if rule == 'hold':
            held = cells[0] if j < 0 else cells[-1]
            return self._maybe_qualify(held, var_id, ctx.current_sheet)
        return self._ship_value(node, ctx)

    def _frame(self, ctx: RenderCtx) -> _Frame:
        """The value the walk is writing an element of: the intermediate
        it entered (``_entered``), else the cell's own line."""
        frame = ctx.frame
        if isinstance(frame, _Frame):
            return frame
        return _Frame(ctx.self_var, 0, _aligned_key(ctx.self_var))

    def _sheet_key(self, var_id: str) -> Optional[tuple]:
        """The window of the sheet a line is written on (its cells run on
        it), or None where its columns are not dated."""
        book = self.time
        sheet = self.var_to_sheet.get(var_id)
        if book is None or sheet is None:
            return None
        st = book.sheets.get(sheet)
        return None if st is None else st.key

    @staticmethod
    def _index_in(var: Any, window: Optional[tuple], frame: _Frame, i: int) -> int:
        """The element of ``var`` (its values on ``window``) Python read
        for ``frame``'s element ``i``: of the same date where ``frame``
        was aligned by date, of the same key where the frame runs along
        the same keys in another order, else of the same position."""
        if frame.window is not None:
            if window is None or window[0] != frame.window[0]:
                return i
            return i - _periods_from(frame.window, window)
        keys = getattr(frame.var, '_keys', None)
        own = getattr(var, '_keys', None)
        if (keys and own and list(own) != list(keys) and 0 <= i < len(keys)
                and keys[i] in own):
            return list(own).index(keys[i])
        return i

    def _read_index(self, var: Any, var_id: str, addr_obj: Any,
                    ctx: RenderCtx) -> Optional[tuple[int, bool]]:
        """``(j, dated)``: the cell of ``var``'s row Python read for the
        element being written, and whether by date (``j`` then may fall
        outside the row: its ``extend=`` rule continues it); None where
        the row's cells are not its own periods, read as before."""
        frame = self._frame(ctx)
        i = ctx.period_idx + frame.offset
        if frame.window is not None:
            window = self._sheet_key(var_id)
            if (window is not None and window[0] == frame.window[0]
                    and (window[2] is None or len(addr_obj.values) == window[2])):
                return self._index_in(var, window, frame, i), True
        if not self._is_native_row(var, addr_obj):
            return None
        return self._index_in(var, None, frame, i), False

    def _entered(self, var: Any, ctx: RenderCtx) -> contextlib.AbstractContextManager[None]:
        """While an unnamed intermediate's expression is written inline,
        the walk writes ``var``'s element — the one Python read for the
        enclosing element (``_index_in``). Where neither it nor the cell
        reads anything but the column's own cells, there is no frame to
        enter."""
        if var is None or (ctx.frame is None and _by_position(ctx.self_var)
                           and _by_position(var)):
            return contextlib.nullcontext()
        return self._entered_frame(var, ctx)

    @contextlib.contextmanager
    def _entered_frame(self, var: Any, ctx: RenderCtx) -> Iterator[None]:
        parent = self._frame(ctx)
        window = _aligned_key(var)
        index = self._index_in(var, window, parent, ctx.period_idx + parent.offset)
        previous = ctx.frame
        ctx.frame = _Frame(var, index - ctx.period_idx, window)
        try:
            yield
        finally:
            ctx.frame = previous

    def _resolve_var_addr(
        self, var_id: str, period_idx: int, current_sheet: Optional[str] = None
    ) -> str:
        addr_obj = self.addresses.get(var_id)
        if addr_obj is None:
            return var_id
        if len(addr_obj.values) == 1:
            addr = addr_obj.values[0]
        elif period_idx < len(addr_obj.values):
            addr = addr_obj.values[period_idx]
        else:
            addr = addr_obj.values[0]
        return self._maybe_qualify(addr, var_id, current_sheet)

    def _maybe_qualify(self, addr: str, var_id: str, current_sheet: Optional[str]) -> str:
        if var_id in self.var_to_sheet:
            var_sheet = self.var_to_sheet[var_id]
            if current_sheet is None or var_sheet != current_sheet:
                return self._format_sheet_reference(var_sheet, addr)
        return addr

    #: Range-taking functions whose Excel signature is VARIADIC, so an
    #: explicit list of cells reads the same as a range. ``IRR`` and
    #: ``XIRR`` are not among them: their second argument is a guess /
    #: a date range, so a comma list would be read as another argument
    #: entirely.
    _LIST_SAFE_FUNCS = {"SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "NPV"}

    #: The variadic functions that read several references as they read
    #: one range: a row whose cells are not one run is written as one
    #: argument per run, ``SUM(Data!C4:E4, Data!G4:I4)``.
    _RUN_FUNCS = {
        "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT", "MEDIAN",
        "STDEV", "STDEVP", "STDEV_S", "STDEV_P", "VAR", "VARP", "VAR_S", "VAR_P", "NPV",
    }

    def _cells_as_range(
        self,
        cells: List[str],
        var_id: str,
        current_sheet: Optional[str],
        ctx: RenderCtx,
        func: Optional[str],
        call: Expr,
    ) -> str:
        # A row's cells are NOT always contiguous: declare quarter
        # totals and the bucket columns stand between the months, so
        # ``SUM(C3:I3)`` swallows the Q1 total and double-counts it —
        # 90 where the model says 60, in a file that looks right.
        # ``totals.rule_formula`` follows the same rule; this is where
        # every range is built, for a whole row and for a slice of one.
        spans = self._runs(cells)
        if len(spans) == 1:
            first = self._maybe_qualify(cells[0], var_id, current_sheet)
            return f"{first}:{cells[-1]}"
        line = (self.var_to_sheet.get(var_id), var_id, tuple(cells))
        if func in self._RUN_FUNCS:
            # One argument per run, each naming its sheet: the call's one
            # range argument becomes as many as the row has runs.
            args = len(spans) + (len(call.args) - 1 if isinstance(call, FuncCall) else 0)
            if args > _MAX_ARGS:
                return self._ship_value(call, ctx, NotOneRun("arguments", (line,), args))
            return ", ".join(self._run_text(cells, lo, hi, var_id, current_sheet)
                             for lo, hi in spans)
        # IRR and the other readers of one ordered range would read a
        # list as further arguments. The truth beats a formula that
        # computes something else: the cell carries its computed value.
        return self._ship_value(call, ctx, NotOneRun(None, (line,)))

    def _format_sheet_reference(self, sheet_name: str, cell_addr: str) -> str:
        if _UNQUOTED_SHEET_NAME.match(sheet_name):
            return f"{sheet_name}!{cell_addr}"
        # Excel doubles an apostrophe inside a quoted name: 'Bob''s'!B2.
        escaped = sheet_name.replace("'", "''")
        return f"'{escaped}'!{cell_addr}"
