# SPDX-License-Identifier: Apache-2.0
"""Excel renderer for the AST walker.

Implements :class:`~modeleon.compile.renderer.Renderer` by rendering each
AST node to an Excel formula fragment. The :class:`ExcelTranslator`
(in ``translator.py``) is the thin facade that wires this renderer up
to a :class:`~modeleon.compile.walker.Walker` and adds the
Excel-specific postprocessing (``=`` prefix, outer-paren stripping).

All Excel specifics live here: operator translation, literal
formatting, cell-reference syntax, range notation, sheet qualification,
function name casing. Adding a new renderer (Google Sheets, JSON, a web
grid) means writing a parallel file alongside this one — core,
translator-facade, and walker stay untouched.
"""

from __future__ import annotations

import logging
import re
import warnings
from datetime import date, datetime
from typing import Any, Dict, List, Optional

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
    RollingAggregate,
    SelfRef,
    Subscript,
    UnaryOp,
    VarRef,
)
from ..renderer import RenderCtx

logger = logging.getLogger(__name__)


# Sheet names that are simple identifiers (starts with letter/underscore,
# rest word chars) can appear unquoted in a reference; anything else
# (spaces, dashes, punctuation, leading digit) needs single-quote wrapping.
_UNQUOTED_SHEET_NAME = re.compile(r"[A-Za-z_]\w*$")


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


# Functions that accept an Excel range for any list-valued VarRef argument.
# The translator renders those args as ranges (``B2:F2``) and leaves
# scalar VarRefs / literals as normal. Covers classic aggregates (SUM,
# MAX, MIN, …) and financial-series funcs (IRR, NPV, XIRR) where Excel
# expects a range for the cash-flow / dates argument.
_RANGE_TAKING_FUNCS = {
    "SUM", "AVERAGE", "MIN", "MAX", "COUNT",
    "IRR", "NPV", "XIRR",
}


# Function names that pass through to Excel unchanged. AST FuncCall nodes
# use Excel casing directly (``IF``, ``ABS``, ``ROUND``, etc.), so the
# translator just emits them verbatim — no rename map needed.
_PASSTHROUGH_FUNCS = _RANGE_TAKING_FUNCS | {
    "IF", "AND", "OR", "NOT", "CHOOSE", "ABS", "ROUND", "INT", "MOD",
    "EDATE", "EOMONTH", "YEAR", "MONTH", "DAY", "DATE", "DAYS360", "TODAY",
    "LEN", "UPPER", "LOWER", "TEXT",  # CONCAT renders as the `&` operator
    "PMT", "FV", "PV", "RATE",
}


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
    ``2025-01-01`` would be read by Excel as arithmetic); numbers pass
    through via ``str()``."""
    if value is None:
        return "0"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return f"DATE({value.year}, {value.month}, {value.day})"
    if isinstance(value, str):
        escaped = value.replace('"', '""')
        return f'"{escaped}"'
    return str(value)


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
    ) -> None:
        self.addresses = addresses
        self.var_to_sheet = var_to_sheet or {}
        self.walker: Any = None  # set by Walker(renderer) constructor

    # ─── Dispatch targets ───────────────────────────────────────

    def render_paren(self, node: Paren, ctx: RenderCtx) -> str:
        """Emit parens only when the wrapped subexpression is compound.

        ``Paren`` in the AST is a precedence-preservation hint, not a
        semantic requirement. When the inner node renders as an atomic
        token (cell reference, literal, function call), the parens are
        redundant — ``(C1) / C3`` carries no more meaning than ``C1 / C3``.
        """
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
        return _value_to_excel_literal(node.value)

    def render_varref(self, node: VarRef, ctx: RenderCtx) -> str:
        var_id = node.var.id
        if var_id in self.addresses:
            return self._resolve_var_addr(var_id, ctx.period_idx, ctx.current_sheet)
        # Variable not in this emission's layout. Two inlining strategies:
        # - if it has its own AST, recurse into it so we reach real cells;
        # - otherwise (plain input referenced from outside the emission
        #   subtree), inline its value as an Excel literal.
        if node.var._expr is not None:
            return self.walker.render(node.var._expr, ctx)
        _warn_out_of_scope_ref(node.var, ctx)
        return _value_to_excel_literal(_value_at_period(node.var._value, ctx.period_idx))

    def render_binop(self, node: BinOp, ctx: RenderCtx) -> str:
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

        left = self._render_operand(node.left, op, ctx, side="left")
        right = self._render_operand(node.right, op, ctx, side="right")
        excel_op = _PY_TO_EXCEL_OP.get(op, op)
        return f"{left} {excel_op} {right}"

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
        rendered = self.walker.render(child, ctx)
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
            logger.warning(
                "Subscript on non-variable expression (%s[%r]) has no Excel "
                "translation — emitting textual fallback.",
                type(node.base).__name__, node.key,
            )
            return f"{self.walker.render(node.base, ctx)}[{node.key}]"

        var_id = node.base.var.id
        addr_obj = self.addresses.get(var_id)
        if addr_obj is None:
            return self._inline_subscript_value(node.base.var, node.key)

        key = node.key
        if isinstance(key, slice):
            # PERIOD-AWARE element, not a range. A slice in an element-
            # wise formula (``cogs = pl_cogs[:20]`` mapping a 26-period
            # driver onto a 20-period statement) must emit the CURRENT
            # period's cell within the sliced window; emitting the whole
            # range stamped the same ``=Input!B1:U1`` into every cell —
            # #SPILL! chaos in Excel 365, accidental implicit
            # intersection in legacy. Aggregate args that genuinely want
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

        if isinstance(key, int) and 0 <= key < len(addr_obj.values):
            return self._maybe_qualify(addr_obj.values[key], var_id, ctx.current_sheet)

        logger.warning(
            "Cannot resolve subscript %r on variable %r — emitting textual "
            "fallback that Excel cannot evaluate. Check that the key exists "
            "in the variable's keys list or is within its value range.",
            key, var_id,
        )
        if isinstance(key, str):
            return f'{var_id}["{key}"]'
        return f"{var_id}[{key}]"

    def render_methodcall(self, node: MethodCall, ctx: RenderCtx) -> str:
        if node.method == "copy":
            return self.walker.render(node.base, ctx)
        if node.method == "shift":
            return self._render_shift(node, ctx)

        logger.warning(
            "Method %r on %r has no Excel translation — emitting textual "
            "fallback that Excel cannot evaluate.",
            node.method, type(node.base).__name__,
        )
        base = self.walker.render(node.base, ctx)
        parts: List[str] = []
        for a in node.args:
            parts.append(self.walker.render(a, ctx) if isinstance(a, Expr) else str(a))
        for k, v in node.kwargs.items():
            v_str = self.walker.render(v, ctx) if isinstance(v, Expr) else str(v)
            parts.append(f"{k}={v_str}")
        return f"{base}.{node.method}({', '.join(parts)})"

    def render_funccall(self, node: FuncCall, ctx: RenderCtx) -> str:
        func = node.func.upper()

        # Backend-specific function: FuncCall declared ``render_backends``
        # without including Excel. Can't render natively — inline the
        # owning Variable's value as a literal, the same fallback used
        # by VarRef when its target is outside the current emission.
        if node.render_backends is not None and 'excel' not in node.render_backends:
            return self._fallback_funccall(node, ctx)

        if func in _RANGE_TAKING_FUNCS:
            parts = [self._render_range_or_cell(a, ctx) for a in node.args]
            return f"{func}({', '.join(parts)})"

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

        if func in _PASSTHROUGH_FUNCS:
            return f"{func}({', '.join(parts)})"

        logger.warning(
            "Unknown function %r in formula — emitting as-is. If this is a "
            "valid Excel function, add it to _PASSTHROUGH_FUNCS; otherwise "
            "the cell will be invalid.",
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
                return _value_to_excel_literal(value)
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
            expansions[name] = self.walker.render(var_expr, ctx)

        if ctx.self_address is None or ctx.period_idx - 1 >= len(ctx.self_address.values):
            prev_str = "0"
        else:
            prev_str = ctx.self_address.values[ctx.period_idx - 1]
        expansions["prev"] = prev_str

        # Support both forms the user may write:
        #   ``"{prev} * (1 + {growth})"``  — explicit braces
        #   ``"prev * (1 + growth)"``       — bare identifiers (Python-style)
        # Python evaluation handles bare names natively (AST eval against the
        # variables dict); without this branch, Excel emission only saw the
        # brace form and left bare names like ``prev`` / ``growth`` un-
        # substituted (literal ``=prev * (1 + growth)`` in the cell).
        def _sub_braced(match: "re.Match[str]") -> str:
            name = match.group(1)
            return expansions.get(name, match.group(0))

        # Unicode identifiers — ``{доход}`` is as legal as ``{growth}``
        # (the ASCII-only class silently skipped Cyrillic placeholders
        # and the bare-name pass then substituted INSIDE the braces).
        emitted = re.sub(r"\{([^\W\d]\w*)\}", _sub_braced, node.template)

        # Substitute bare-identifier occurrences with word boundaries so
        # ``growth`` doesn't accidentally rewrite a longer identifier like
        # ``growth_rate``. Replace longest names first so a shorter name
        # (``g``) doesn't shadow a longer one (``growth``).
        for name in sorted(expansions, key=len, reverse=True):
            emitted = re.sub(
                rf"\b{re.escape(name)}\b", expansions[name], emitted
            )

        return emitted

    def render_rollingaggregate(self, node: RollingAggregate, ctx: RenderCtx) -> str:
        if ctx.period_idx < node.window - 1:
            return _value_to_excel_literal(node.fill)

        var_id = node.source.id
        addr_obj = self.addresses.get(var_id)
        if addr_obj is None or not addr_obj.values:
            return _value_to_excel_literal(node.fill)

        start_idx = ctx.period_idx - node.window + 1
        end_idx = min(ctx.period_idx, len(addr_obj.values) - 1)
        first = self._maybe_qualify(addr_obj.values[start_idx], var_id, ctx.current_sheet)
        last = addr_obj.values[end_idx]
        return f"{node.func}({first}:{last})"

    def render_regrain(self, node: Regrain, ctx: RenderCtx) -> str:
        i = ctx.period_idx
        fill = node.fill_values[i] if i < len(node.fill_values) else 0.0
        var_id = node.source.id
        addr_obj = self.addresses.get(var_id)
        if (addr_obj is None or not addr_obj.values or i >= len(node.buckets)):
            return _value_to_excel_literal(fill)
        lo, hi = node.buckets[i]
        n = len(addr_obj.values)
        if lo >= n or hi - 1 >= n:
            return _value_to_excel_literal(fill)
        if node.recipe == 'first':
            return self._maybe_qualify(addr_obj.values[lo], var_id, ctx.current_sheet)
        if node.recipe == 'last':
            return self._maybe_qualify(addr_obj.values[hi - 1], var_id, ctx.current_sheet)
        fn = {'sum': 'SUM', 'mean': 'AVERAGE', 'min': 'MIN', 'max': 'MAX'}.get(node.recipe)
        if fn is None:
            return _value_to_excel_literal(fill)        # geometric etc. -> inlined value
        first = self._maybe_qualify(addr_obj.values[lo], var_id, ctx.current_sheet)
        last = addr_obj.values[hi - 1]
        return f"{fn}({first}:{last})"

    # ─── Excel-specific helpers ─────────────────────────────────

    def _needs_parens(self, inner: Expr) -> bool:
        if isinstance(inner, (BinOp, UnaryOp, Compare)):
            return True
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

    def _render_shift(self, node: MethodCall, ctx: RenderCtx) -> str:
        if not isinstance(node.base, VarRef):
            return f"{self.walker.render(node.base, ctx)}.shift({node.args[0]})"

        var_id = node.base.var.id
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
            if isinstance(fill_value, Expr):
                return self.walker.render(fill_value, ctx)
            return _value_to_excel_literal(fill_value)

        addr_obj = self.addresses.get(var_id)
        if addr_obj is None:
            return _fill()
        if target_period < 0 or target_period >= len(addr_obj.values):
            return _fill()

        return self._maybe_qualify(addr_obj.values[target_period], var_id, ctx.current_sheet)

    def _render_range_or_cell(self, arg: Expr, ctx: RenderCtx) -> str:
        if isinstance(arg, VarRef):
            var_id = arg.var.id
            addr_obj = self.addresses.get(var_id)
            if addr_obj is not None and len(addr_obj.values) > 1:
                return self._resolve_var_range(var_id, ctx.current_sheet)
        # A SLICED row inside an aggregate wants the sliced RANGE:
        # ``SUM(x[0:12])`` → ``SUM(B1:M1)``. (Element-wise slice
        # rendering lives in ``render_subscript`` and is period-aware;
        # this is the one context where the whole window is the point.)
        # The slice arrives either as a bare Subscript node or — the
        # common DSL shape — as a VarRef to the intermediate Variable
        # ``x[0:12]`` produced (its ``_expr`` is the Subscript).
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
            var_id = sub.base.var.id
            addr_obj = self.addresses.get(var_id)
            if addr_obj is not None:
                values = addr_obj.values[sub.key]
                if len(values) > 1:
                    first = self._maybe_qualify(
                        values[0], var_id, ctx.current_sheet
                    )
                    return f"{first}:{values[-1]}"
        return self.walker.render(arg, ctx)

    def _inline_subscript_value(self, base_var: Any, key: Any) -> str:
        value = base_var._value
        keys = getattr(base_var, "_keys", None)

        if isinstance(key, slice):
            if isinstance(value, list):
                return "[" + ", ".join(_value_to_excel_literal(v) for v in value[key]) + "]"
            return (base_var.python_name or base_var.path.leaf)

        if isinstance(key, str) and keys and key in keys:
            return _value_to_excel_literal(value[keys.index(key)])
        if (
            isinstance(key, int)
            and keys
            and key in keys
            and key not in range(len(value) if isinstance(value, list) else 1)
        ):
            return _value_to_excel_literal(value[keys.index(key)])
        if isinstance(key, int) and isinstance(value, list) and 0 <= key < len(value):
            return _value_to_excel_literal(value[key])

        var_id = (base_var.python_name or base_var.path.leaf)
        if isinstance(key, str):
            return f'{var_id}["{key}"]'
        return f"{var_id}[{key}]"

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

    def _resolve_var_range(self, var_id: str, current_sheet: Optional[str] = None) -> str:
        addr_obj = self.addresses.get(var_id)
        if addr_obj is None:
            return var_id
        if len(addr_obj.values) == 1:
            return self._maybe_qualify(addr_obj.values[0], var_id, current_sheet)
        first = self._maybe_qualify(addr_obj.values[0], var_id, current_sheet)
        last = addr_obj.values[-1]
        return f"{first}:{last}"

    def _format_sheet_reference(self, sheet_name: str, cell_addr: str) -> str:
        if _UNQUOTED_SHEET_NAME.match(sheet_name):
            return f"{sheet_name}!{cell_addr}"
        return f"'{sheet_name}'!{cell_addr}"
