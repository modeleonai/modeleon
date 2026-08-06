# SPDX-License-Identifier: Apache-2.0
"""Expression tree — structured representation of Variable formulas.

Each ``Variable`` that holds a formula owns an ``Expr`` describing it. The
tree is the source of truth; the ``Variable.formula`` string is derived
from ``Expr.to_string()``.

Node types cover every formula pattern the DSL currently produces:

- Atoms: :class:`Literal`, :class:`VarRef`
- Arithmetic / comparison: :class:`BinOp`, :class:`UnaryOp`, :class:`Compare`
- Container access: :class:`Subscript`, :class:`ListExpr`
- Function-like: :class:`FuncCall`, :class:`MethodCall`
- Roll-forward: :class:`SelfRef`

Every node implements:

- ``to_string()`` — render back to a readable formula string. Consumers
  of ``Variable.formula`` (JSON exports, HTML repr, debug) get this
  form.
- ``iter_refs()`` — yield every ``Variable`` referenced anywhere in the
  subtree. Used for dependency analysis and cycle detection.

Tree-walking translators (Python formula → Excel formula) live in
:mod:`modeleon.compile.excel.translator`.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Iterator, List, Optional, TYPE_CHECKING, Union

if TYPE_CHECKING:
    from .variable import Variable


# ─── Base ────────────────────────────────────────────────────────────


class Expr:
    """Abstract base for all expression-tree nodes."""

    def to_string(self) -> str:
        """Render this node as the formula string stored in Variable.formula."""
        raise NotImplementedError

    def iter_refs(self) -> Iterator["Variable"]:
        """Yield every ``Variable`` this subtree depends on, in traversal order.

        May yield duplicates if the same Variable is referenced in multiple
        places. Callers that need uniqueness should deduplicate (by identity).
        """
        raise NotImplementedError


# ─── Atoms ──────────────────────────────────────────────────────────


@dataclass(frozen=True)
class Literal(Expr):
    """A constant scalar — number, string, bool."""

    value: Any

    def to_string(self) -> str:
        if isinstance(self.value, str):
            # Python-style quoted (``'hello'``) so ``.formula`` round-trips
            # cleanly. The translator renders the same Literal with Excel-
            # style double-quotes for the actual cell.
            return repr(self.value)
        return str(self.value)

    def iter_refs(self) -> Iterator["Variable"]:
        return iter(())


@dataclass
class VarRef(Expr):
    """Reference to an existing ``Variable``.

    Holds a direct Python reference — name lookups (``.id``) happen at
    render time, so renaming a Variable propagates automatically into any
    existing formulas.
    """

    var: "Variable"

    def to_string(self) -> str:
        # Prefer python_name (set by ``parent.child = ...``). For unattached
        # Variables with an explicit display_name, snake-case the label so
        # formulas read ``revenue * 0.6`` instead of ``v3 * 0.6``.
        if self.var.python_name:
            return self.var.python_name
        explicit_label = getattr(self.var, '_display_name', None)
        if explicit_label:
            from .humanize import identifier_from_label
            return identifier_from_label(explicit_label)
        # Anonymous intermediate (Variable produced by operator overloading
        # — ``a * b`` between adopted Variables creates one that has its own
        # ``_expr`` but never gets a ``python_name``). Emitting its floating
        # ``path.leaf`` here would surface as ``vffff7ab483b0`` in rendered
        # formulas. Inline the expression instead so chains like
        # ``(revenue * cogs_pct) * units`` render as their composite
        # algebra rather than an opaque floating id. Wrap in parens for
        # precedence safety — the inlined expression may itself be a BinOp.
        inner_expr = getattr(self.var, '_expr', None)
        if inner_expr is not None:
            return f"({inner_expr.to_string()})"
        return self.var.path.leaf

    def iter_refs(self) -> Iterator["Variable"]:
        yield self.var


# ─── Grouping ───────────────────────────────────────────────────────


@dataclass
class Paren(Expr):
    """Explicit ``(inner)`` grouping — preserves precedence in rendered form.

    Used by operators that historically wrapped compound operands to keep the
    formula string unambiguous (e.g. division: ``(a + b) / c``). Nodes don't
    imply precedence on their own — the renderer emits them as-is.
    """

    inner: Expr

    def to_string(self) -> str:
        return f"({self.inner.to_string()})"

    def iter_refs(self) -> Iterator["Variable"]:
        yield from self.inner.iter_refs()


# ─── Arithmetic and comparison ──────────────────────────────────────


@dataclass
class BinOp(Expr):
    """Binary arithmetic: ``left OP right`` for +, -, *, /, //, %, **."""

    op: str  # one of: + - * / // % **
    left: Expr
    right: Expr

    def to_string(self) -> str:
        return f"{self.left.to_string()} {self.op} {self.right.to_string()}"

    def iter_refs(self) -> Iterator["Variable"]:
        yield from self.left.iter_refs()
        yield from self.right.iter_refs()


@dataclass
class UnaryOp(Expr):
    """Unary prefix operator. Only ``-`` is used today, but keep extensible."""

    op: str  # "-"
    operand: Expr

    def to_string(self) -> str:
        return f"{self.op}{self.operand.to_string()}"

    def iter_refs(self) -> Iterator["Variable"]:
        yield from self.operand.iter_refs()


@dataclass
class Compare(Expr):
    """Binary comparison: ``left OP right`` for ==, !=, <, <=, >, >=."""

    op: str  # one of: == != < <= > >=
    left: Expr
    right: Expr

    def to_string(self) -> str:
        return f"{self.left.to_string()} {self.op} {self.right.to_string()}"

    def iter_refs(self) -> Iterator["Variable"]:
        yield from self.left.iter_refs()
        yield from self.right.iter_refs()


# ─── Subscript / list ──────────────────────────────────────────────


@dataclass
class Subscript(Expr):
    """``base[key]`` where key may be int, str, or slice.

    - ``revenue[0]`` — positional, period or list index.
    - ``scenarios["Bear"]`` — named key from ``_keys``.
    - ``revenue[1:4]`` — slice; currently preserved as text but not supported
      beyond display.
    """

    base: Expr
    key: Union[int, str, slice]

    def to_string(self) -> str:
        base_s = self.base.to_string()
        if isinstance(self.key, slice):
            start = "" if self.key.start is None else self.key.start
            stop = "" if self.key.stop is None else self.key.stop
            if self.key.step is not None:
                return f"{base_s}[{start}:{stop}:{self.key.step}]"
            return f"{base_s}[{start}:{stop}]"
        if isinstance(self.key, str):
            return f'{base_s}["{self.key}"]'
        return f"{base_s}[{self.key}]"

    def iter_refs(self) -> Iterator["Variable"]:
        yield from self.base.iter_refs()


@dataclass
class ListExpr(Expr):
    """``[a, 0, b, 5]`` — compound formula with mixed Variables and scalars.

    Produced by :func:`Variable._build_compound_formula` for list-mode
    construction. Each item is either a ``VarRef`` or a ``Literal``.
    """

    items: List[Expr] = field(default_factory=list)

    def to_string(self) -> str:
        return f"[{', '.join(item.to_string() for item in self.items)}]"

    def iter_refs(self) -> Iterator["Variable"]:
        for item in self.items:
            yield from item.iter_refs()


# ─── Function and method calls ─────────────────────────────────────


@dataclass
class FuncCall(Expr):
    """Free-function call: ``IF(cond, a, b)``, ``sum(x)``, ``max(a, b, c)``.

    Used for DSL helpers that aren't methods on a specific Variable.

    Two optional compatibility hints:

    - ``render_backends`` — renderers that can emit native target output
      for this function. ``None`` (default) means universal; every
      renderer handles it. A set like ``frozenset({'excel'})`` declares
      the function render-native only to that target; other renderers
      fall back to inlining ``_value`` as a literal.
    - ``compute_backends`` — compute environments that can evaluate this
      function natively at runtime. ``None`` (default) means universal
      end-to-end. Populate only when render and compute diverge
      (e.g. a SQL renderer can emit ``IRR(col)`` but the SQL engine
      needs a UDF to actually evaluate it).
    """

    func: str
    args: List[Expr] = field(default_factory=list)
    render_backends: Optional[frozenset] = None
    compute_backends: Optional[frozenset] = None

    def to_string(self) -> str:
        return f"{self.func}({', '.join(a.to_string() for a in self.args)})"

    def iter_refs(self) -> Iterator["Variable"]:
        for arg in self.args:
            yield from arg.iter_refs()


@dataclass
class MethodCall(Expr):
    """Method call on a Variable: ``revenue.shift(1)``, ``.pct_change(1)``, ``.copy()``.

    ``args`` and ``kwargs`` may contain ``Expr`` nodes (rare) or plain
    Python values (common — shift periods, fill_value, etc.).
    """

    base: Expr
    method: str
    args: List[Any] = field(default_factory=list)
    kwargs: Dict[str, Any] = field(default_factory=dict)

    def to_string(self) -> str:
        parts: List[str] = []
        for a in self.args:
            parts.append(a.to_string() if isinstance(a, Expr) else str(a))
        for k, v in self.kwargs.items():
            v_str = v.to_string() if isinstance(v, Expr) else str(v)
            parts.append(f"{k}={v_str}")
        return f"{self.base.to_string()}.{self.method}({', '.join(parts)})"

    def iter_refs(self) -> Iterator["Variable"]:
        yield from self.base.iter_refs()
        for a in self.args:
            if isinstance(a, Expr):
                yield from a.iter_refs()
        for v in self.kwargs.values():
            if isinstance(v, Expr):
                yield from v.iter_refs()


# ─── Roll-forward ──────────────────────────────────────────────────


@dataclass
class SelfRef(Expr):
    """``recurrence(start, template, **variables, periods=...)`` roll-forward.

    The template string (e.g. ``"{prev} + {flow}"``) stays a string here —
    parsing user-supplied templates into a sub-AST is a separate project and
    not needed to fix the self-reference translation bug. What matters for
    the translator is that :attr:`variables` resolves placeholders to
    concrete Variable references, and :attr:`start` is structured so the
    translator can emit the correct period-0 cell.
    """

    start: Expr
    template: str
    variables: Dict[str, Expr] = field(default_factory=dict)
    periods: Optional[Union[int, Expr]] = None

    def to_string(self) -> str:
        parts = [self.start.to_string(), f'"{self.template}"']
        for name, var_expr in self.variables.items():
            parts.append(f"{name}={var_expr.to_string()}")
        if self.periods is not None:
            if isinstance(self.periods, Expr):
                parts.append(f"periods={self.periods.to_string()}")
            else:
                parts.append(f"periods={self.periods}")
        return f"recurrence({', '.join(parts)})"

    def iter_refs(self) -> Iterator["Variable"]:
        yield from self.start.iter_refs()
        for var_expr in self.variables.values():
            yield from var_expr.iter_refs()
        if isinstance(self.periods, Expr):
            yield from self.periods.iter_refs()


# ─── Rolling-window aggregate ──────────────────────────────────────


@dataclass
class RollingAggregate(Expr):
    """Per-period rolling aggregate over a source Variable's cell range.

    At period ``t``, emits ``=FUNC(source[t-window+1]:source[t])``, giving
    the correct per-cell range reference for ``rolling_sum`` /
    ``rolling_mean`` / any other windowed aggregate Excel supports natively.
    Cells before the window fills (``t < window - 1``) emit the literal
    ``fill``, matching the Python-side semantics.
    """

    source: "Variable"
    window: int
    func: str            # "SUM", "AVERAGE", etc.
    fill: Any = 0.0      # literal returned before the window is full

    def to_string(self) -> str:
        label = self.source.python_name or self.source.path.leaf
        return f"{label}.rolling_{self.func.lower()}({self.window})"

    def iter_refs(self) -> Iterator["Variable"]:
        yield self.source


# ─── Re-grain (coarser-grain projection of a source) ───────────────


@dataclass
class Restrict(Expr):
    """Drop-one-axis coordinate restrict — ``var.at(track='actual')``.

    The canonical slice of a Variable with tracks (§17.2): one coordinate of
    one quantity, as its own auditable node. ``axis`` is the axis name
    ('tracks' — v1's only finite axis); ``label`` is the ROLE (the
    machine key). Renderers without a native lowering fall back to the inlined
    track value; the Excel lowering (a reference into the coordinate's
    laid-out row) arrives with the P3 surface work.
    """

    base: Expr
    label: str
    axis: str = 'tracks'

    def to_string(self) -> str:
        return f"{self.base.to_string()}.at({self.axis}={self.label!r})"

    def iter_refs(self) -> Iterator["Variable"]:
        yield from self.base.iter_refs()


@dataclass
class Regrain(Expr):
    """Per-bucket re-grain of a source Variable onto a coarser grain.

    At target period ``i``, emits a reducer over the source's native cells in
    bucket ``i`` — ``=SUM(source[lo]:source[hi-1])`` for a flow, the period-end
    cell for a stock, etc. The source's native cells must be laid out for the
    reference to resolve; if the source has no address in the workbook the
    renderer falls back to the inlined value (``fill_values[i]``).
    """

    source: "Variable"
    buckets: List[Any]      # list of (lo, hi) half-open native-index ranges
    recipe: str             # 'sum' | 'mean' | 'first' | 'last' | 'min' | 'max' | ...
    fill_values: List[Any]  # eager re-grained values, used when source has no address

    def to_string(self) -> str:
        label = self.source.python_name or self.source.path.leaf
        return f"{label}.regrain({self.recipe})"

    def iter_refs(self) -> Iterator["Variable"]:
        yield self.source
