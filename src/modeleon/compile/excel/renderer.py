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
    "IF", "ABS", "ROUND", "INT", "MOD",
    "EDATE", "EOMONTH", "YEAR", "MONTH", "DAY", "DATE", "TODAY",
    "LEN", "UPPER", "LOWER", "CONCAT", "TEXT",
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
    is emitted elsewhere); numbers pass through via ``str()``."""
    if value is None:
        return "0"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
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
        return "[" + ", ".join(self.walker.render(item, ctx) for item in node.items) + "]"

    def render_literal(self, node: Literal, ctx: RenderCtx) -> str:
        return _value_to_excel_literal(node.value)

    def render_varref(self, node: VarRef, ctx: RenderCtx) -> str:
        var_id = node.var.id
        if var_id in self.addresses:
            return self._resolve_var_addr(var_id, ctx.period_idx, ctx.current_sheet)
        # Variable not in this emission's layout. Two inlining strategies:
        # - if it has its own AST, recurse into it so we reach real cells;
        # - otherwise (plain input referenced from outside the emission
        #   subtree), Policy A: inline its value as an Excel literal.
        if node.var._expr is not None:
            return self.walker.render(node.var._expr, ctx)
        _warn_out_of_scope_ref(node.var, ctx)
        return _value_to_excel_literal(_value_at_period(node.var._value, ctx.period_idx))

    def render_binop(self, node: BinOp, ctx: RenderCtx) -> str:
        left = self.walker.render(node.left, ctx)
        right = self.walker.render(node.right, ctx)
        op = node.op
        # // → INT(a/b), % → MOD(a,b) — Excel has no direct operator for these
        if op == "//":
            return f"INT({left}/{right})"
        if op == "%":
            return f"MOD({left},{right})"
        excel_op = _PY_TO_EXCEL_OP.get(op, op)
        return f"{left} {excel_op} {right}"

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
            values = addr_obj.values[key]
            if not values:
                return ""
            if len(values) == 1:
                return self._maybe_qualify(values[0], var_id, ctx.current_sheet)
            first = self._maybe_qualify(values[0], var_id, ctx.current_sheet)
            return f"{first}:{values[-1]}"

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
        # owning Variable's value as a literal (Policy A extended from
        # VarRef).
        if node.render_backends is not None and 'excel' not in node.render_backends:
            return self._fallback_funccall(node, ctx)

        if func in _RANGE_TAKING_FUNCS:
            parts = [self._render_range_or_cell(a, ctx) for a in node.args]
            return f"{func}({', '.join(parts)})"

        parts = [self.walker.render(a, ctx) for a in node.args]
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

        def _sub(match: "re.Match[str]") -> str:
            name = match.group(1)
            return expansions.get(name, match.group(0))

        return re.sub(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", _sub, node.template)

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

        fill_value = node.kwargs.get("fill_value", "0")
        target_period = ctx.period_idx - periods_int

        addr_obj = self.addresses.get(var_id)
        if addr_obj is None:
            return str(fill_value)
        if target_period < 0 or target_period >= len(addr_obj.values):
            return str(fill_value)

        return self._maybe_qualify(addr_obj.values[target_period], var_id, ctx.current_sheet)

    def _render_range_or_cell(self, arg: Expr, ctx: RenderCtx) -> str:
        if isinstance(arg, VarRef):
            var_id = arg.var.id
            addr_obj = self.addresses.get(var_id)
            if addr_obj is not None and len(addr_obj.values) > 1:
                return self._resolve_var_range(var_id, ctx.current_sheet)
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
