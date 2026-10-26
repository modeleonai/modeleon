# SPDX-License-Identifier: Apache-2.0
"""Variable — the atomic cell primitive.

A Variable holds either a literal value or a formula. Operator
overloading (``*``, ``+``, ``-``, ``/``, …) on Variables eagerly computes
results AND records each intermediate as an :class:`Expr` AST node, so
the writer can render live Excel formulas at output time instead of
dead values.

This file owns identity, construction entry point, metadata, container
protocols, and representation. Parallel files handle specific concerns:

- :mod:`modeleon.core.variable_init`  — construction dispatch
  (value / formula / pyformula → ``_value`` / ``_expr`` / deps)
- :mod:`modeleon.core.variable_ops`   — arithmetic + comparison
  operators, broadcasting, unit propagation
- :mod:`modeleon.core.expr`           — AST node types used to build
  ``self._expr`` trees
- :mod:`modeleon.core.qpath`          — ``QPath`` identity type; the
  floating-namespace path for unrooted nodes is minted from
  ``id(self)`` inline (no shared counter)
- :mod:`modeleon.core.errors`         — :class:`CircularDependencyError`

Variable inherits from two mixins: :class:`_VariableInit` (construction
dispatchers) and :class:`_VariableArithmetic` (operators). The class
declaration below wires them in; callers don't touch the mixins
directly.
"""

from collections import namedtuple
from contextvars import ContextVar
from typing import (Optional, Set, Any, Dict, List, Sequence, Tuple, TYPE_CHECKING, Self,
                    overload)

from .humanize import humanize_identifier
from .component import Component
from .expr import Expr, MethodCall, Paren, Regrain, Restrict, Subscript, TimeRef, VarRef
from .qpath import QPath
from .regrain import RegrainSpec
from .shape import Shape
from .tracks import TrackValues
from .variable_init import _VariableInit
from .variable_ops import _VariableArithmetic, _forget_placements

if TYPE_CHECKING:
    from pathlib import Path

    from .unit import Unit

    from .loops import Dependents



#: A Variable's native time location — ``(start, grain)`` — or ``None`` when
#: time-agnostic. Returned by :attr:`Variable.time`. ``first_day`` is the
#: day a short first period starts on (a quarter / year window started
#: inside its first period), ``None`` when the first period is whole.
TimeLoc = namedtuple('TimeLoc', ['start', 'grain', 'first_day'], defaults=(None,))


def is_typed_date(var: object) -> bool:
    """``mo.DATE(2026, 1, 1)`` of three whole numbers: a date the author
    typed. As a line of its own it is an input, as
    ``mo.Variable(date(2026, 1, 1))`` is; inside a formula it stays the
    ``DATE(...)`` call it reads as."""
    from .expr import FuncCall, Literal
    expr = getattr(var, '_expr', None)
    return (isinstance(expr, FuncCall) and expr.func == 'DATE' and len(expr.args) == 3
            and all(isinstance(a, Literal) and type(a.value) is int for a in expr.args))


def _TrackValuesType() -> type:
    from .tracks import TrackValues
    return TrackValues


def _typed_dates_line(var: "Variable") -> bool:
    """A line whose formula is only typed dates: one ``mo.DATE`` of
    numbers, or a list of them (``[mo.DATE(2026, 1, 1), ...]``) beside
    plain values - never a reference to another line."""
    from .expr import ListExpr, Literal, VarRef
    if is_typed_date(var):
        return True
    expr = getattr(var, '_expr', None)
    if not isinstance(expr, ListExpr):
        return False
    typed = 0
    for item in expr.items:
        if isinstance(item, VarRef) and item.var._owner is None and is_typed_date(item.var):
            typed += 1
        elif not isinstance(item, Literal):
            return False
    return typed > 0



def series_at_grain(label: str, values: Any, loc: 'TimeLoc', spec: Any, grain: str,
                    blank: Any) -> tuple:
    """The rule half of :meth:`Variable.at` for a LOCATED series: ``(values,
    start_label, first_day, ranges, rule)`` at ``grain`` — or ``at``'s refusal,
    word for word. One body for the method and for the projection's track-wise
    fold (``core.projection``), so a tracked line's roles re-grain exactly as
    the line would and a refusal names the line the same way."""
    from .regrain import Ratio
    from .time import bucket_ranges, coarse_first_day, regrain_series
    if spec is None:
        raise ValueError(
            f"{label!r} has no re-grain rule; give it one (e.g. "
            f"regrain=mo.up('sum'/'last'/'mean')) to re-grain {loc.grain} "
            f"-> {grain}."
        )
    if spec.frozen and spec.default is None and not spec.overrides:
        raise ValueError(
            f"{label!r} is grain-frozen with no value rule — it has no "
            f"meaning off its native grain (a period counter / date "
            f"spine). If its VALUES do have coarse meaning — quarter-end "
            f"cash is the last monthly close — declare it: "
            f"regrain=mo.frozen(values='last')."
        )
    # A frozen spec with a value rule: the formula never re-evaluates at
    # the target grain, but the computed value series re-grains below
    # like any stored series — which is exactly what this path does.
    rule = spec.overrides.get((loc.grain, grain), spec.default)
    if rule is None:
        raise ValueError(f"{label!r} has no re-grain rule for {loc.grain} -> {grain}.")
    if isinstance(rule, Ratio) or callable(rule):
        raise ValueError(
            "ratio and callable re-grain rules are not applied by "
            "Variable.at() — a ratio rule re-grains through the model "
            "projection (model.at(grain)); use a named recipe here."
        )
    n = len(values) if isinstance(values, list) else 1
    ranges = bucket_ranges(loc.grain, grain, loc.start, n)
    labels, out = regrain_series(values, loc.grain, grain, loc.start, rule, loc.first_day,
                                 blank=blank)
    return out, labels[0], coarse_first_day(loc.start, loc.grain, loc.first_day, grain), ranges, rule


#: Every field ``Variable.__init__`` starts from, with its default — written
#: into a new instance's ``__dict__`` in one step; the constructor then
#: writes only what differs. Immutable values only (the two per-instance
#: containers, ``_excel_props`` and ``_dependency_refs``, are built per
#: instance); the key set is exactly the fields the constructor defines,
#: so ``vars(v)`` / ``hasattr`` see every field. ``_source_code`` is not
#: among them: it is derived on read (:class:`_SourceCode`) and a value in
#: ``__dict__`` would shadow the descriptor.
_VARIABLE_DEFAULTS: Dict[str, Any] = {
    '_role_kwargs': None, '_roles_materialized': False, '_roles_provisional': False,
    '_blend_name': None, '_excel_layout': None, '_owner': None, '_component_name': None,
    '_value': None, '_expr': None, '_raw_formula_str': None, 'pyformula': None,
    '_unit': None, '_unit_declared': False, '_keys': None,
    '_keys_source': None, 'var_type': 'scalar', 'value_type': 'float', '_indexed_by': (),
    '_start': None, '_grain': None, '_regrain': None, '_schedule': None, '_extend': None,
}


#: The compounds whose shared text is being decided in this context. A row of
#: a loop container that lists itself (``row = Variable([row, 2.0])``) is
#: rewritten in place, so its items reach the row again; such a list shares
#: no text.
_SOURCE_RESOLVING: ContextVar[frozenset] = ContextVar(
    'modeleon_source_resolving', default=frozenset()
)


def _shared_source_code(owner: 'Variable', items: Tuple['Variable', ...]) -> Optional[str]:
    """The one source text every Variable in ``items`` carries, else None."""
    resolving = _SOURCE_RESOLVING.get()
    if id(owner) in resolving:
        return None
    token = _SOURCE_RESOLVING.set(resolving | {id(owner)})
    try:
        shared: Optional[str] = None
        for item in items:
            code = item._source_code
            if not code or (shared is not None and code != shared):
                return None
            shared = code
        return shared
    finally:
        _SOURCE_RESOLVING.reset(token)


class _SourceCode:
    """``Variable._source_code``, derived when read rather than at construction.

    A non-data descriptor (``__get__`` only): a text written on the
    instance — a helper's resolved call (``PMT(0.005, 360, ...)``), the
    public setter, the ``None`` an adopted typed-dates line writes — lands
    in ``__dict__`` and wins. With none written, the text follows what the
    instance says it is: its own formula (``_source_is_formula``, the IF
    family) or the one text its list items share (``_source_items``, a
    ``Variable([a, b, ...])`` compound), else None. The formula render
    inlines every anonymous operand, so the helpers used to pay for it at
    every call for a text almost nobody reads; nested conditions paid for
    the inner text twice.
    """

    @overload
    def __get__(self, obj: None, objtype: Optional[type] = None) -> '_SourceCode': ...

    @overload
    def __get__(self, obj: 'Variable', objtype: Optional[type] = None) -> Optional[str]: ...

    def __get__(self, obj: Optional['Variable'], objtype: Optional[type] = None) -> Any:
        if obj is None:
            return self
        if obj._source_is_formula:
            return obj.formula
        items = obj._source_items
        if items is not None:
            return _shared_source_code(obj, items)
        return None

_object_setattr = object.__setattr__

def _under_settings(node: Any) -> bool:
    """True when ``node`` is, or sits inside, a container that holds
    settings rather than lines (an ExcelView)."""
    seen: set = set()
    while node is not None and id(node) not in seen:
        seen.add(id(node))
        if getattr(type(node), '_settings_only', False):
            return True
        node = getattr(node, '__dict__', {}).get('_parent')
    return False


class Variable(_VariableInit, _VariableArithmetic, Component):
    # ``__eq__`` is overridden on ``_VariableArithmetic`` to return a
    # comparison Variable (formula building), not a bool. Python drops
    # the default ``__hash__`` when ``__eq__`` is overridden — restore
    # identity-based hashing so Variables can be members of sets,
    # WeakSets, and dict keys.
    __hash__ = object.__hash__

    # Forward references (see :mod:`modeleon.core.loops`): a placeholder
    # for a row read before it is assigned carries ``(owner, name)`` and
    # the Variables built on it; a Variable computed from a placeholder
    # carries the set of placeholders it waits on.
    _forward_of: Optional[Tuple[Any, str]] = None
    # The day a short first period starts on, for a row located at its own
    # window (``mo.time`` rows, re-grained series); see ``TimeLoc``.
    _first_day: Any = None
    _forward_dependents: Optional['Dependents'] = None
    _forward_horizon: Optional[int] = None
    _forward_site: Optional[Tuple[str, int]] = None
    _forward_abandoned: bool = False
    _awaits: Optional[Set['Variable']] = None
    _loop_error: Optional[str] = None
    #: Does this value depend on ``mo.time`` (itself, or through the
    #: unnamed intermediates it is built from)? Set with the expression.
    _reads_time: bool = False
    #: Were its operands read by DATE onto its window (``_start``) — two
    #: windows met in an operator or a function that is one call per
    #: period? A renderer then reads each operand's cell of the same date;
    #: a value on a window of its own read otherwise (``start=``, a
    #: re-grain, a copy) keeps its operands' positions.
    _read_by_date: bool = False
    #: Where ``_source_code`` comes from when no text was written on the
    #: instance (see :class:`_SourceCode`): the Variable's own formula, or
    #: the text the Variables of its list share.
    _source_is_formula: bool = False
    _source_items: Optional[Tuple['Variable', ...]] = None
    _source_code = _SourceCode()

    """
    Pure variable - computation logic only

    Example (scalar input variable):
        revenue = Variable(1000000)  # Value first! Name injected from metadata

    Example (scalar computed variable):
        cogs = revenue * 0.65  # is_formula is automatically True with operators
        gross_profit = revenue - cogs

    Example (list):
        years = Variable([2025, 2026, 2027, 2028, 2029], var_type="list", value_type="int")

    Example (with explicit display label):
        revenue = Variable(1_000_000, display_name="Total Revenue")
    """
    
    # Valid keys inside ``excel_props``. Passing any other key raises.
    # Mirrors ``MultiVariableBase._EXCEL_PROP_KEYS`` — kept separate here
    # so the two classes can diverge if variable-specific props appear.
    _EXCEL_PROP_KEYS = frozenset({
        # Typography
        'bold', 'italic', 'font_color', 'font_size', 'font_family',
        'text_align', 'indent',
        # Fill / borders
        'bg', 'border_top', 'border_bottom', 'border_left', 'border_right',
        # Value-driven fill: paint a value cell's background this color
        # wherever the cell holds a nonzero number — a Gantt band / active
        # mask reads as a continuous bar. Emitted as a LIVE conditional
        # rule in the workbook, so edited values repaint themselves.
        'bg_nonzero',
        # Row-total opt-in: when the sheet's view projects a
        # ``('row_sum', caption)`` lead column, this row prints its
        # periods' total there — a LIVE ``=SUM(...)`` in the workbook,
        # the computed number in any other renderer. Rows without the
        # flag leave the column empty (a ratio row's sum is nonsense).
        'row_sum',
        # Utility row: laid out and computed as usual (values,
        # formulas, labels) but NOT SHOWN — a natively hidden row in
        # the .xlsx; other renderers skip it. Inherited down the MV tree.
        'hidden',
        # Data formatting
        'number_format',
        # Layout: this row's VALUES print on its parent section's
        # header row — the financial-statement subtotal-on-header form.
        'header_row',
        # The row's article number ("1.1") — data, not label text; the
        # view decides whether it renders as a lead column or inline.
        'article',
        # The line is its block's RESULT — what a calculation block
        # computes, as opposed to the steps and links above it. A mark,
        # not a look: the view's ``bands['result']`` says how it shows.
        'result',
        # Cell-type colouring — cascades to descendants. ``True`` for the
        # default palette, or a dict ``{'input'|'formula'|'reference': hex}``.
        'format_by_type',
    })

    @classmethod
    def to_new_source(cls, name: str) -> str:
        """Source form of a NEW, empty instance bound to ``name`` — the
        class describes its own constructor spelling (the source sibling
        of ``__repr__``, mirroring
        :meth:`~modeleon.core.multi_variable.MultiVariableBase.to_new_source`).
        The name binds on the LHS, so the spelling ignores it; a fresh
        variable starts as the zero scalar.
        """
        return "mo.Variable(0)"

    def __init__(self, value=None, value_type: str = "float", var_type: str = "scalar",
                 display_name: Optional[str] = None, description: Optional[str] = None,
                 formula=None, pyformula=None, unit=None, keys=None,
                 excel_props=None, excel_layout: Optional[Any] = None,
                 indexed_by: Sequence['Variable'] = (),
                 start: Optional[str] = None, grain: Optional[str] = None,
                 regrain: Optional['RegrainSpec'] = None,
                 extend: Optional[Any] = None,
                 tracks: Optional[dict] = None,
                 **kwargs):
        """Create a Variable from exactly one of ``value``, ``formula``, or ``pyformula``.

        Args:
            value: Literal input data. Scalar, list, dict (becomes keyed list),
                date/datetime, or another Variable (copies it).
            value_type: Data type (``"float"``, ``"int"``, ``"string"``,
                ``"datetime"``). Auto-detected from ``value`` when possible.
            var_type: Shape (``"scalar"`` or ``"list"``). Auto-detected too.
            display_name: Human-readable label shown in the spreadsheet
                header. Defaults to the Python identifier discovered via
                introspection (``revenue = mo.Variable(1000)`` picks up
                ``"revenue"``), then to the auto id.
            formula: Formula string (``"revenue * 0.65"``), a Variable to copy,
                or a list of Variables / scalars (compound formula).
            pyformula: Value produced by arbitrary Python (e.g.
                ``pyformula=my_func(v1, v2)``). Treated as an *input* — the
                resulting Variable is marked ``is_formula=False`` — but the
                string rendering of the passed value is captured on
                ``self.pyformula`` so downstream tools can regenerate the
                producing expression.
            keys: Named index for list values — list of labels or another
                Variable whose values become keys.
            excel_props: Dict of Excel-renderer hints — styling
                (``'bold'``, ``'bg'``, ``'number_format'``, …) and
                formatting. See :attr:`_EXCEL_PROP_KEYS` for valid keys.
                Unknown keys raise ``TypeError``.
            excel_layout: Optional opaque layout-config object. Engine
                stores it as ``self._excel_layout`` and reads back the
                ``format`` attribute when writing this Variable's cells
                (other fields ignored at the Variable layer — they're
                meaningful at the MV layer for sheet-level
                decisions). Any object with a ``format`` attribute
                works; the engine stays duck-typed.
            indexed_by: Sequence of axis Variables this Variable is
                laid out along. An axis is just another Variable —
                typically one with a list of labels. Operator-built
                Variables get this set automatically by broadcast; pure
                inputs declare it here. Default ``()`` means scalar.
            **kwargs: Plugin-registered kwargs (e.g. ``tag=`` from an
                extension). Unknown kwargs raise TypeError so typos
                don't silently disappear.

        ``value``, ``formula``, and ``pyformula`` are mutually exclusive.
        For arithmetic-produced Variables (``total = a + b``) the formula is
        built by the operator methods and ``__init__`` is called with
        ``formula=<string>``.
        """
        # The fields every Variable starts from, in one C-level write
        # (``_VARIABLE_DEFAULTS``): the body below writes a field only when
        # its value is NOT the default. Per-field writes through
        # ``__setattr__`` dominated the cost of a Variable(), and every
        # operator in a formula builds one. The key set of ``__dict__`` is
        # the same either way.
        self.__dict__.update(_VARIABLE_DEFAULTS)

        # Dict sugar: Variable({"Bear": 800, "Base": 1000}) → list+keys
        if isinstance(value, dict):
            keys = list(value.keys())
            value = list(value.values())

        if (value is not None) + (formula is not None) + (pyformula is not None) > 1:
            raise ValueError(
                "Cannot specify more than one of 'value', 'formula', or 'pyformula'. "
                "Use 'value' for literal values, 'formula' for calculations."
            )

        # --- track-kwarg authoring ---
        # ``mo.Variable(actual=[...], budget=<expr>)`` — the coordinate
        # kwargs. Track names are USER CONTENT (declared once on the
        # model via ``tracks=mo.Tracks(...)``), so they cannot be
        # literal parameters; the stash is filled from the leftover
        # **kwargs after plugin consumption (see the ``_process_kwargs``
        # call below) and validated against the ambient declaration at
        # ADOPTION — the constructor only records intent.
        # The override spelling: a positional shared expression plus
        # track overrides — ``mo.Variable(price * seats, actual=ledger)``
        # — derives non-overridden tracks from the shared value.
        # ``_role_kwargs`` (None until kwargs arrive), ``_roles_materialized``
        # (True once adoption has lowered the track kwargs onto a
        # TrackValues — a FLAG, not a check on the value's type: a tracked
        # SHARED expression makes ``_value`` TrackValues at construction,
        # and the overrides still owe a materialization pass),
        # ``_roles_provisional`` (True while the value was materialized
        # from the KWARGS' OWN names, before any declaration was
        # reachable; adoption validates and clears it) and
        # ``_blend_name`` (the track the BLEND synthesized here, None until
        # synthesis — tells synthesized-by-us from authored on re-runs)
        # start from the defaults above.

        # --- tracked construction (the tracks value) ---
        # ``tracks={'plan': [...], 'actual': [...]}`` builds the value as
        # role → time-series. Engine-internal spelling the role kwargs
        # lower onto at adoption. Mutually exclusive with the flat forms
        # — a Variable is tracked or flat, never both.
        if tracks is not None:
            if value is not None or formula is not None or pyformula is not None:
                raise ValueError(
                    "tracks= is exclusive with value/formula/pyformula — "
                    "a coordinate's own expression binds inside the "
                    "tracks dict."
                )
            if indexed_by:
                raise ValueError(
                    "tracks= IS the finite-axis value — indexed_by= "
                    "declares a separate axis and cannot be combined "
                    "with it (axis budget is one)."
                )
            from .tracks import TrackValues
            value = TrackValues(tracks) if not isinstance(tracks, TrackValues) else tracks

        # --- identity + default fields ---
        # ``_qualified_id``, ``_python_name``, ``_display_name``, and
        # ``_excel_props`` are initialized on the :class:`Component`
        # parent (with excel_props validated against ``_EXCEL_PROP_KEYS``).
        # The ``.path`` property below synthesizes a floating-namespace
        # path from ``id(self)`` until adoption crystallizes a rooted
        # path.
        if (display_name is None and excel_props is None and description is None
                and type(self) is Variable):
            # What ``Component.__init__`` (with ``Base.__init__``) writes
            # when given nothing, in its order: empty identity fields and a
            # fresh props dict. Every operator result is built this way, and
            # walking that chain — a ``__setattr__`` call per field — was
            # over half the cost of a ``Variable()``. Keep the two in step
            # (``test_component.py`` compares them). A subclass may extend
            # the chain, so it always walks it.
            fields = self.__dict__
            fields['_qualified_id'] = None
            fields['_python_name'] = None
            fields['_display_name'] = None
            fields['_description'] = None
            fields['_excel_props'] = {}
        else:
            super().__init__(
                display_name=display_name,
                excel_props=excel_props,
                description=description,
            )
        # Opaque layout-config attribute. Writer reads ``.format`` off it
        # to apply per-cell number_format overrides. ``None`` is the
        # inherit-from-parent signal for a resolver that walks the tree.
        if excel_layout is not None:
            self._excel_layout = excel_layout
        # Adoption back-pointers ``_owner`` / ``_component_name`` are set by
        # ``MultiVariableBase._register_component`` when this Variable is
        # attached to a parent MV — None until then, so the ``.path``
        # property reads them without ``getattr`` fallbacks. ``_value``,
        # ``_expr`` (the structured source of truth for the formula; the
        # ``.formula`` string property is a read-through view of
        # ``_expr.to_string()`` — set ``_expr``, not ``formula``),
        # ``_raw_formula_str`` (the fallback string form for formulas
        # passed or assigned as raw strings), ``pyformula`` (the captured
        # rendering of a ``pyformula=`` input, which marks the Variable as
        # an input) all start as None; ``_source_code`` (a friendlier source
        # snippet the ``functions/*`` helpers set) is derived on read.
        # _dependency_refs mirrors _expr.iter_refs() (with dedup); populated by
        # construction sites that build an AST. Kept as a cached list rather
        # than a live property so callers don't re-walk the tree on every read.
        self._dependency_refs: List['Variable'] = []
        if unit is not None:
            resolved_unit = self._resolve_unit(unit)
            if resolved_unit is not None:
                self._unit = resolved_unit
                # A unit the author declares (``unit=`` / ``set_unit``) is kept
                # as given; one the arithmetic derives may be derived again (a
                # loop's rows, built on placeholders, are).
                self._unit_declared = True
        if keys is not None:
            self._keys, self._keys_source = self._resolve_keys(keys)

        # --- dispatch to the mode that was requested ---
        if formula is not None:
            var_type = self._init_from_formula(formula, var_type)
        elif pyformula is not None:
            var_type = self._init_from_pyformula(pyformula, var_type)
        elif value is not None:
            var_type, value_type = self._init_from_value(value, var_type, value_type)
        # else: empty variable — all defaults apply

        # --- validate keys vs values length ---
        if self._keys is not None and self._value is not None:
            val_len = len(self._value) if isinstance(self._value, list) else 1
            if len(self._keys) != val_len:
                raise ValueError(
                    f"Mismatched keys/values: {len(self._keys)} keys "
                    f"({self._keys!r}) but {val_len} value(s). "
                    f"Each key labels exactly one value — pass matching lengths, "
                    f"or use the dict form ``Variable({{'label': value, ...}})`` "
                    f"which keeps them in sync automatically."
                )
            if not isinstance(self._value, list):
                self._value = [self._value]
                var_type = 'list'

        # --- final type + metadata ---
        if var_type != 'scalar':
            self.var_type = var_type
        if value_type != 'float':
            self.value_type = value_type
        # is_formula is a derived @property — see below.
        # ``_display_name`` and ``_excel_props`` were set on Component
        # via ``super().__init__`` above.

        # --- shape via dep graph ---
        # ``_indexed_by`` is the authoritative declaration of which axes
        # this Variable is laid out along. Operator-built Variables get
        # this set automatically by ``_VariableArithmetic`` from the
        # union of operand axes. Pure inputs declare it explicitly via
        # the ``indexed_by=`` constructor kwarg.
        #
        # The dep graph itself (``_dependency_refs``) carries the rest
        # of the shape story: axes are just Variables that other
        # Variables happen to reference via ``indexed_by``. There is
        # no separate AxisProvider registry.
        if indexed_by:
            self._indexed_by = tuple(indexed_by)

        # --- axised-construction invariant ---
        # A declared shape must be honored by the value, or the axis
        # label is a lie the first consumer trips over. Flat row-major
        # only: the canonical cell order is the left-first axis order,
        # last axis fastest — nested lists are ambiguous about which
        # nesting is which axis, so they are rejected outright.
        if self._indexed_by and self._value is not None:
            flat = self._value if isinstance(self._value, list) else [self._value]
            if any(isinstance(x, list) for x in flat):
                raise ValueError(
                    "indexed_by= values must be a FLAT row-major list "
                    "(left-first axis order, last axis fastest), not "
                    "nested lists — nesting is ambiguous about which "
                    "level is which axis. Flatten the values: for shape "
                    "(A×B) pass [a0b0, a0b1, a1b0, a1b1, ...]."
                )
            from .shape import Shape
            expected = Shape(self._indexed_by).length
            if len(flat) != expected:
                axis_names = ", ".join(
                    str(getattr(a, '_display_name', None)
                        or getattr(a, '_python_name', None) or 'axis')
                    + f"({len(a._value) if isinstance(a._value, list) else 1})"
                    for a in self._indexed_by
                )
                raise ValueError(
                    f"indexed_by shape mismatch: declared axes [{axis_names}] "
                    f"require {expected} value(s), got {len(flat)}. Pass one "
                    f"value per cell of the declared shape, flat row-major."
                )

        # --- ambient time location + re-grain ---
        # A Variable is DEFINED at a native (start, grain) — re-grainable
        # metadata — or is time-agnostic (no grain; broadcasts unchanged).
        if start is not None:
            self._start = start
        if grain is not None:
            self._grain = grain
        if regrain is not None:
            self._regrain = regrain
        # ``mo.schedule`` steps ({date: value}, materialized at adoption
        # when the ambient window resolves) and the declared out-of-extent
        # behavior (mo.zero()/mo.hold()/mo.none()) — zero/hold
        # materialize a partial series to the window at adoption;
        # arithmetic between unequal extents aligns them by date and
        # reads each side's out-of-extent cells through its rule.
        # ``_schedule`` starts as None.
        if extend is not None:
            from .extend import ExtendRule
            if not isinstance(extend, ExtendRule):
                raise TypeError(
                    "extend= takes mo.zero(), mo.hold() or mo.none() — the "
                    "declared continuation of a partial series beyond its "
                    "values."
                )
            self._extend = extend.kind
        # A positional Variable carries its PENDING materialization
        # state into the copy: ``mo.Variable(mo.schedule({...}),
        # actual=...)`` binds the schedule Variable before it ever adopts
        # (its ``_value`` is still None), so the wrapper inherits the
        # steps / extension rule and materializes them itself.
        if isinstance(value, Variable):
            if getattr(value, '_schedule', None) is not None:
                self._schedule = dict(value._schedule)
                # The steps travel with their re-grain rule: a wrapped
                # schedule is the same rate as the bare one, so it
                # averages into a quarter the same way unless the
                # wrapper declares a rule of its own.
                if self._regrain is None and value._regrain is not None:
                    self._regrain = value._regrain
            if self._extend is None:
                inherited_extend = getattr(value, '_extend', None)
                if inherited_extend is not None:
                    self._extend = inherited_extend
            # Wrapping never strips units: ``mo.Variable(a + b,
            # display_name=…)`` keeps the sum's unit unless the wrapper
            # declares its own.
            if self._unit is None and getattr(value, '_unit', None) is not None:
                self._unit = value._unit
                if getattr(value, '_unit_declared', False):
                    self._unit_declared = True
            # Nor its keys or axes: ``mo.Variable(growth * 2, display_name=…)``
            # still runs along the keys the product ran along.
            wrapped_keys = getattr(value, '_keys', None)
            if self._keys is None and wrapped_keys:
                self._keys = list(wrapped_keys)
                self._keys_source = getattr(value, '_keys_source', None)
            if not self._indexed_by and getattr(value, '_indexed_by', None):
                self._indexed_by = tuple(value._indexed_by)
        if (start is None) != (grain is None):
            raise ValueError(
                "A located Variable needs both `start` and `grain` (or neither, "
                "for a time-agnostic value)."
            )
        if grain is not None:
            from .time import GRAINS
            if grain not in GRAINS:
                raise ValueError(
                    f"Unknown time grain {grain!r}; supported: {', '.join(GRAINS)}."
                )
        if regrain is not None:
            if not isinstance(regrain, RegrainSpec):
                raise TypeError(
                    "`regrain=` takes a RegrainSpec — use mo.up(...), "
                    "mo.frozen(), or mo.ratio(...)."
                )
            # Re-grain applies each line's rule to its OWN computed value, so a
            # FORMULA carries a rule too (a flow sums its own series; a ratio
            # uses mo.ratio). grain may be inherited from an ancestor MV's
            # default_grain (resolved lazily by ``.time``), so it is not required
            # here — a rule with no resolvable grain is caught at projection.

        if kwargs:
            _track_candidates = self._process_kwargs(kwargs)
            if _track_candidates:
                if tracks is not None:
                    raise ValueError(
                        "track kwargs and tracks= are two spellings of one "
                        "thing — pick one."
                    )
                if formula is not None or pyformula is not None:
                    raise ValueError(
                        "track kwargs replace formula=/pyformula= — bind "
                        "the shared expression positionally: "
                        "mo.Variable(<shared formula>, actual=[...])."
                    )
                if indexed_by:
                    raise ValueError(
                        "track kwargs ARE the finite axis — indexed_by= "
                        "cannot be combined with them (axis budget is one)."
                    )
                self._role_kwargs = _track_candidates
                # Materialize PROVISIONALLY from the kwargs' own names.
                # The declaration lives on the model, which a Variable
                # built inside a MultiVariableClass body cannot reach —
                # and the very next line (``self.revenue = self.hours *
                # rate``) computes eagerly, at operator time. The kwargs already
                # NAME their tracks, so the value is knowable here;
                # adoption validates the names against the declaration,
                # fills any track the shared expression owes, and attaches
                # the blend. Skipped when a shared positional expression is
                # present — the non-overridden tracks need the declared
                # vocabulary, which only adoption has.
                if value is None and formula is None and pyformula is None:
                    self._materialize_roles(
                        list(_track_candidates), display_name or 'variable'
                    )
                    self._roles_provisional = self._roles_materialized

    @property
    def shape(self) -> Shape:
        """Tuple of axes this Variable is laid out along.

        Derived on every access from ``_indexed_by`` — pure dep-graph
        read, no separate registry. Operator-built Variables have
        ``_indexed_by`` populated by the broadcast in
        :class:`~modeleon.core.variable_ops._VariableArithmetic`; pure
        inputs declared their axes via the ``indexed_by=`` kwarg.

        Scalars and Variables with no axis declarations return
        ``Shape(())``.
        """
        return Shape(self._indexed_by)

    @property
    def time(self) -> Optional['TimeLoc']:
        """The Variable's effective time location ``(start, grain)``, or ``None``
        when it is time-agnostic.

        Resolves the Variable's own ``(start, grain)`` first; if it declares
        none, it inherits ``default_start`` / ``default_grain`` from the nearest
        ancestor MV that sets them — the same ``_owner`` → ``_parent`` cascade
        ``resolve_excel_view`` uses. No ancestor default → agnostic (``None``).

        AXISED Variables have no single time location (the rank-1 guard):
        the flat value enumerates COORDINATES, not periods — claiming an
        inherited monthly window over it would let a grain projection sum
        scenario cells into quarters with no error. A series per
        coordinate is spelled with tracks, where time lives inside each
        track; an ``indexed_by`` variable stays time-less.
        """
        if self._indexed_by:
            return None
        if self._grain is not None:
            return TimeLoc(self._start, self._grain, self._first_day)
        # Only list-valued (time-series) Variables inherit an ancestor's
        # default grain; a scalar constant is time-agnostic — it broadcasts
        # unchanged into any projection — regardless of a model default.
        # A FIBERED value with list tracks is a time series too — time
        # lives INSIDE each track (tracks outside, periods inside).
        if isinstance(self._value, TrackValues):
            if self._value.time_length is None:
                return None
        elif not isinstance(self._value, list):
            return None
        seen: Set[int] = set()
        node = getattr(self, '_owner', None) or getattr(self, '_parent', None)
        while node is not None and id(node) not in seen:
            seen.add(id(node))
            # A container's window is set on the instance; a container
            # without one is asked without its component fallback.
            slots = node.__dict__ if hasattr(node, '__dict__') else {}
            grain = slots.get('default_grain')
            if grain is None and 'default_grain' not in slots:
                grain = getattr(type(node), 'default_grain', None)
                if not isinstance(grain, str):
                    grain = None
            if grain is not None:
                return TimeLoc(getattr(node, 'default_start', None), grain,
                               slots.get('_default_first_day'))
            node = getattr(node, '_owner', None) or getattr(node, '_parent', None)
        return None

    def _on_adopted(self) -> None:
        """Adoption hook — runs when this Variable joins an MV tree
        (``_register_component``) and again at crystallization (when a
        floating subtree is mounted under a rooted parent).

        Two duties, both needing the ambient window that only the
        ancestor chain can resolve:

        1. **Materialize a schedule** (``mo.schedule``): map each window
           period to the step in force at its date.
        2. **The extent law**: a bare list must be length 1 or
           exactly the window length. A partial series is legal only when
           declared partial (own ``start=``, a schedule, or ``extend=``).
           Without it, a 3-element list in a 24-period model is accepted
           silently and detonates at an unrelated line.

        Before either, the formulas placed through this line while it had
        no window (``variable_ops._forget_placements``) are placed again:
        it runs on a window now, or on its container's once that has one.

        Idempotent; a no-op while the window is unresolvable (floating
        subtrees get their check at crystallization).
        """
        from .time import resolve_default_window, section_window_refusal, _period_labels
        owner = getattr(self, '_owner', None) or getattr(self, '_parent', None)
        if owner is None:
            return
        if '_placed_through' in self.__dict__:
            _forget_placements(self)
        if _under_settings(owner):
            # A view's settings (``timeline={'totals': ['quarter', 'year']}``)
            # are words, not a series of the window: their length is not
            # checked against the window.
            return
        if not isinstance(self._value, _TrackValuesType()) and _typed_dates_line(self):
            # ``inputs.close = mo.DATE(2026, 1, 1)``: the line is a date
            # typed in, written to the book as its value - not a formula.
            self.expr = None
            self._source_code = None
            self.value_type = 'datetime'
        # Only a series that takes its periods from the section is placed
        # by its window: a scalar, or a line located on its own (an
        # ``.at()`` copy, ``start=``), takes nothing from it.
        from .tracks import TrackValues as _TrackValues
        _v = self._value
        if getattr(self, '_grain', None) is None and (
                isinstance(_v, list)
                or (isinstance(_v, _TrackValues) and _v.time_length is not None)):
            refusal = section_window_refusal(owner, reads_time=False)
            if refusal is not None:
                label = (self._display_name or getattr(self, '_python_name', None)
                         or 'A line')
                raise ValueError(f"{label!r} joins a section whose window is not whole: {refusal}")
        window = resolve_default_window(owner)
        # Scalar track coordinates need no window (time lives INSIDE a
        # track; an all-scalar TrackValues is time-agnostic), so the
        # window gate scopes the TIME duties below — it must not swallow
        # role materialization on a windowless model.
        have_window = window is not None and window.grain is not None

        from .tracks import TrackValues
        materialized_here = False
        if (have_window and self._schedule is not None
                and not isinstance(self._value, (list, TrackValues))):
            if window.start is None or window.periods is None:
                raise ValueError(
                    "mo.schedule needs a fully declared window — set "
                    "default_start and default_periods on the model."
                )
            labels = _period_labels(
                window.start, None, window.periods, window.grain
            )
            steps = self._schedule
            step_dates = list(steps.keys())
            label_set = set(labels)
            for d in step_dates:
                # Later steps may legitimately lie beyond the window; a
                # step INSIDE the window span must land on a boundary.
                if labels[0] <= d <= labels[-1] and d not in label_set:
                    raise ValueError(
                        f"mo.schedule step {d!r} does not land on a "
                        f"{window.grain} boundary of the window "
                        f"({labels[0]} … {labels[-1]}) — move the step "
                        f"to a period start, or change the model grain."
                    )
            if step_dates[0] > labels[0]:
                raise ValueError(
                    f"mo.schedule starts at {step_dates[0]!r} but the "
                    f"window starts at {labels[0]!r} — add the value in "
                    f"force at the window start."
                )
            values = []
            current = None
            i = 0
            for lb in labels:
                while i < len(step_dates) and step_dates[i] <= lb:
                    current = steps[step_dates[i]]
                    i += 1
                values.append(current)
            self._value = values
            self.var_type = 'list'
            materialized_here = True
            # No return: a schedule can be the SHARED value under role
            # kwargs (override spelling) — materialization continues below.

        # --- declared extension: materialize a partial series ---
        # ``extend=mo.zero()/mo.hold()`` legalizes the prefix spelling:
        # the author SAID what lies beyond the values, so the series
        # completes to the window here — the same moment a schedule
        # materializes. ``none`` stays un-materialized (its meaning is
        # absent/NA, not a value to pad with).
        if (
            have_window
            and self._extend in ('zero', 'hold')
            and isinstance(self._value, list)
            and window.periods is not None
            and 0 < len(self._value) < window.periods
            and self._grain is None
            and not self._indexed_by
            and self._keys is None
        ):
            pad = 0.0 if self._extend == 'zero' else self._value[-1]
            self._value = list(self._value) + (
                [pad] * (window.periods - len(self._value))
            )
            self.var_type = 'list'
            materialized_here = True

        # --- a line that materialized its own values stops re-reading ---
        # Wrapping or cloning a source that still holds the pending form
        # (``mo.Variable(mo.schedule({...}))``, a partial series with
        # ``extend=``, a schedule cloned out of a container with no
        # window) hands this line the steps / extension, and the line
        # materialized them itself just above. Its expression, though,
        # only re-reads the source, which still holds the pending form
        # (no value yet, or the short list), not the numbers this line
        # now carries. Written as a reference, the workbook would show
        # the source's form instead (an absent value inlines as ``=0``).
        # So the line becomes the input it is and writes its own
        # values, exactly like the bare spelling. The same holds for a
        # formula the line's ``extend=`` carried past its operands'
        # periods (``mo.Variable(z * 2, extend=mo.zero())``): no formula
        # says "0 from here on". A source that already holds these very
        # numbers keeps the reference: it is true. So does a formula
        # reading a series placed by its own ``start=`` / ``grain=``:
        # the padding above ignored that placement, and the export must
        # still see the read to refuse it rather than write the series
        # under the wrong dates.
        if materialized_here and self._expr is not None:
            reread = self._expr
            if (isinstance(reread, MethodCall) and reread.method == 'copy'
                    and not reread.args and not reread.kwargs):
                reread = reread.base    # ``source.copy()`` — a clone
            held = reread.var._value if isinstance(reread, VarRef) else None
            placed = any(getattr(v, '_grain', None) is not None
                         for v in self._expr.iter_refs())
            if not placed and not (
                    isinstance(held, list) and held == self._value):
                self.expr = None

        # --- role materialization ---
        # ``mo.Variable(actual=..., budget=...)`` recorded intent at
        # construction; the ambient ``mo.Tracks`` declaration is
        # reachable only from the tree, so the lowering happens here.
        # Idempotence rides the ``_roles_materialized`` FLAG — a tracked
        # shared expression makes ``_value`` TrackValues at construction
        # already, and its overrides still owe this pass.
        # A provisionally-materialized variable still owes this pass:
        # its names came from the kwargs, and only the declaration can
        # say whether they are TRACKS or typos.
        tracks_decl = None
        if self._role_kwargs is not None and (
                not self._roles_materialized or self._roles_provisional):
            from .tracks_decl import resolve_tracks_decl
            label = (self._display_name
                     or getattr(self, '_python_name', None) or 'variable')
            decl = tracks_decl = resolve_tracks_decl(owner)
            if decl is None:
                # The window and the declaration may live on DIFFERENT
                # ancestors: a floating subtree with its own window
                # cannot see the model's mo.Tracks yet. Defer until the
                # chain tops out at a Model — the canonical root; only
                # there is "no declaration" a final answer.
                from .model import Model
                top = owner
                _seen: Set[int] = set()
                while id(top) not in _seen:
                    _seen.add(id(top))
                    nxt = (getattr(top, '_owner', None)
                           or getattr(top, '_parent', None))
                    if nxt is None:
                        break
                    top = nxt
                if not isinstance(top, Model):
                    return    # floating subtree — re-fires at mount
                raise ValueError(
                    f"{label!r} got kwarg(s) "
                    f"{', '.join(f'{r}=' for r in self._role_kwargs)} "
                    f"that are neither Variable options nor declared "
                    f"tracks. If they are tracks, declare the axis once "
                    f"on the model — tracks=mo.Tracks("
                    f"{', '.join(map(repr, self._role_kwargs))}); if "
                    f"not, fix the kwarg name (known: value, formula, "
                    f"pyformula, display_name, unit, keys, excel_props, "
                    f"start, grain, regrain, extend)."
                )
            self._validate_tracks_against(decl, label)
            if not self._roles_materialized:
                self._materialize_roles(decl.names, label)
            self._roles_provisional = False

        if not have_window:
            return    # the extent laws below need the window

        # --- extent law, tracked form: every track obeys the window ---
        if isinstance(self._value, TrackValues):
            tl = self._value.time_length
            if (tl is not None and tl > 1 and window.periods is not None
                    and tl != window.periods and self._grain is None):
                label = (self._display_name
                         or getattr(self, '_python_name', None) or 'variable')
                raise ValueError(
                    f"{label!r}: tracks carry {tl} value(s) in a "
                    f"{window.periods}-period window ({window.grain}) — "
                    f"each track is a time series and must cover the "
                    f"window. Ragged tracks are not supported."
                )
            # --- live synthesis — AFTER the extent law, so a
            # ragged track dies with the teaching error, never a raw
            # IndexError from the splice loop. The declaration resolved
            # above (when this line owed the role pass) still holds:
            # nothing since has touched the containers.
            self._synthesize_blend(owner, window, tracks_decl)
            return

        # --- extent law ---
        if (
            isinstance(self._value, list)
            and len(self._value) > 1
            and window.periods is not None
            and len(self._value) != window.periods
            and self._grain is None          # no own location declared
            and not self._indexed_by         # coordinates, not periods
            and self._keys is None           # labeled list, not a series
            and self._schedule is None
            and self._extend is None
        ):
            label = (self._display_name
                     or getattr(self, '_python_name', None) or 'variable')
            raise ValueError(
                f"{label!r} has {len(self._value)} value(s) in a "
                f"{window.periods}-period window ({window.grain}). A bare "
                f"list must cover the whole window or be a single value. "
                f"For a partial series say what continues it: "
                f"extend=mo.zero() (a flow — absent beyond its values), "
                f"extend=mo.hold() (a rate — the last value stays), "
                f"start='YYYY-MM' for its own anchor, or "
                f"mo.schedule({{...}}) for date-keyed steps."
            )

    # Plain (non-underscore) instance attributes the engine itself
    # assigns — everything else non-underscore that is not a class
    # descriptor is a TRACK BINDING (an MV-style write).
    _PLAIN_ATTRS = frozenset({'var_type', 'value_type', 'pyformula'})

    def __setattr__(self, name: str, value: Any) -> None:
        """Dotted track write — ``revenue.actual = [...]``.

        The MV-style incremental spelling: a late-arriving actual lands
        AFTER the line is defined, as its own statement, without
        rewriting the definition. Internal attributes (underscore),
        engine-assigned plain attributes, and class descriptors
        (properties) take the
        normal path; any other name is a track binding, validated
        against the ambient declaration when the tree can reach it and
        stashed until adoption when it cannot.
        """
        # ``self._PLAIN_ATTRS``, not the class's: a subclass may extend it
        # with plain attributes of its own.
        if name[:1] == '_' or name in self._PLAIN_ATTRS or hasattr(type(self), name):
            _object_setattr(self, name, value)
            return
        self._bind_track(name, value)

    def _bind_track(self, name: str, operand: Any) -> None:
        from .time import resolve_default_window
        from .tracks import TrackValues
        from .tracks_decl import resolve_tracks_decl

        label = (self._display_name
                 or getattr(self, '_python_name', None) or 'variable')
        owner = self._owner or getattr(self, '_parent', None)
        decl = resolve_tracks_decl(owner) if owner is not None else None
        if decl is None:
            # Floating (or undeclared) — record intent; adoption
            # validates exactly like constructor track kwargs.
            stash = dict(self._role_kwargs or {})
            stash[name] = operand
            self._role_kwargs = stash
            self._roles_materialized = False
            return
        if name not in decl:
            if name == getattr(decl.blend, 'name', None):
                raise ValueError(
                    f"{label!r}: {name!r} is the SYNTHESIZED track "
                    f"(mo.blend name={name!r}) — it is derived from "
                    f"{decl.blend.given!r}/{decl.blend.follow!r} and "
                    f"cannot be written directly; bind those instead."
                )
            raise ValueError(
                f"{label!r}: {name!r} is not a declared track "
                f"({', '.join(decl.names)}) — add it to mo.Tracks(...), "
                f"or fix the name if it is a typo."
            )
        # Extract the bound value (same rules as adoption-time
        # materialization: loud on pending/valueless operands,
        # coordinate-aligned pull from tracked ones, identity dep edge).
        v = operand
        if isinstance(v, Variable):
            if (getattr(v, '_role_kwargs', None) is not None
                    and not v._roles_materialized):
                raise ValueError(
                    f"{label!r}.{name}: reads a variable whose own "
                    f"tracks have not materialized yet — attach it to "
                    f"the model first."
                )
            ov = v._value
            if ov is None:
                raise ValueError(
                    f"{label!r}.{name}: reads a variable that has no "
                    f"value yet (still floating, or an unadopted "
                    f"mo.schedule) — attach it to the model first."
                )
            if isinstance(ov, TrackValues):
                if name not in ov:
                    raise ValueError(
                        f"{label!r}.{name}: the operand carries tracks "
                        f"({', '.join(ov.roles)}) but no {name!r} — "
                        f"slice explicitly: .at(track='...')."
                    )
                ov = ov[name]
            if not any(r is v for r in self._dependency_refs):
                self._dependency_refs.append(v)
            v = list(ov) if isinstance(ov, list) else ov
        elif isinstance(v, list):
            v = self._lower_role_cells(list(v), name, label)
        window = resolve_default_window(owner)
        if (isinstance(v, list) and len(v) > 1 and window is not None
                and window.periods is not None
                and len(v) != window.periods):
            raise ValueError(
                f"{label!r}.{name} carries {len(v)} value(s) in a "
                f"{window.periods}-period window — a track binding "
                f"covers the whole window or is a single value."
            )
        cur = self._value
        if isinstance(cur, TrackValues):
            merged = cur.as_dict()
            merged[name] = v
        elif cur is None:
            # Incremental authoring: tracks accumulate one statement at
            # a time; a subset is legal until arithmetic meets the
            # mismatch law.
            merged = {name: v}
        else:
            # A flat shared value expands into every declared track,
            # then the binding overrides one.
            merged = {
                n: (list(cur) if isinstance(cur, list) else cur)
                for n in decl.names
            }
            merged[name] = v
        self._value = TrackValues(merged)
        stash = dict(self._role_kwargs or {})
        stash[name] = operand
        self._role_kwargs = stash
        self._roles_materialized = True
        if self._value.time_length is not None:
            self.var_type = 'list'
        # A binding changes the source data — the synthesized blend
        # track must re-splice from the new tracks.
        if window is not None and window.grain is not None:
            self._synthesize_blend(owner, window, decl)

    def _validate_tracks_against(self, decl, label: str) -> None:
        """Every authored track name must be declared."""
        undeclared = [r for r in (self._role_kwargs or {}) if r not in decl]
        if not undeclared:
            return
        blend_name = getattr(decl.blend, 'name', None)
        if blend_name in undeclared:
            raise ValueError(
                f"{label!r} authors {blend_name!r}, but the "
                f"declaration synthesizes that track (mo.blend "
                f"name={blend_name!r}) — rename the authored "
                f"track, or drop the blend."
            )
        raise ValueError(
            f"{label!r} got kwarg(s) "
            f"{', '.join(map(repr, undeclared))} but the model "
            f"declares tracks {', '.join(decl.names)} — add the "
            f"track to mo.Tracks(...), or fix the kwarg name if "
            f"it is a typo."
        )

    def _lower_role_cells(self, items: list, role: str, label: str) -> list:
        """Unwrap Variable ELEMENTS of a role-bound list.

        The positional path lowers a compound list through
        ``_build_compound_formula``; the role path used to store the
        raw objects, and Variable instances leaked into
        :class:`TrackValues` — and from there into every renderer and
        serializer. Elements lower by the same rules as a whole-Variable
        operand: value pull, role-coordinate alignment for tracked
        elements, loud on valueless ones, identity dependency edge.
        """
        from .tracks import TrackValues
        out: list = []
        for x in items:
            if not isinstance(x, Variable):
                out.append(x)
                continue
            xv = x._value
            if xv is None:
                raise ValueError(
                    f"{label!r}.{role}: a list element reads a variable "
                    f"that has no value yet (still floating, or an "
                    f"unadopted mo.schedule) — attach it to the model "
                    f"first."
                )
            if isinstance(xv, TrackValues):
                if role not in xv:
                    raise ValueError(
                        f"{label!r}.{role}: a list element carries "
                        f"tracks ({', '.join(xv.roles)}) but no "
                        f"{role!r} — slice explicitly: .at(track='...')."
                    )
                xv = xv[role]
            out.append(xv)
            if not any(r is x for r in self._dependency_refs):
                self._dependency_refs.append(x)
        return out

    def _materialize_roles(self, names, label: str) -> None:
        """Lower the stashed track kwargs onto a :class:`TrackValues`.

        ``names`` is the ordered track vocabulary — the model's
        declaration at adoption, or the kwargs' own names when a
        variable materializes PROVISIONALLY inside a class body (where
        no declaration is reachable and the next line's arithmetic
        needs a value).
        """
        from .tracks import TrackValues
        shared = self._value
        if shared is None and self._schedule is not None:
            raise ValueError(
                "mo.schedule needs a fully declared window — set "
                "default_grain, default_start and default_periods "
                "on the model."
            )
        new_tracks: Dict[str, Any] = {}
        for role in names:
            operand = (self._role_kwargs or {}).get(role)
            if operand is None:
                # The override spelling: non-overridden roles
                # derive from the shared positional expression.
                # Without one, the track simply isn't authored yet —
                # a SUBSET is legal (incremental authoring).
                if shared is None:
                    continue
                if isinstance(shared, TrackValues):
                    if role not in shared:
                        raise ValueError(
                            f"{label!r}: the shared expression "
                            f"carries no {role!r} coordinate "
                            f"({', '.join(shared.roles)}) — author "
                            f"{role}= explicitly."
                        )
                    src = shared[role]
                else:
                    src = shared
                new_tracks[role] = (
                    list(src) if isinstance(src, list) else src
                )
                continue
            if isinstance(operand, Variable):
                if (getattr(operand, '_role_kwargs', None) is not None
                        and not operand._roles_materialized):
                    raise ValueError(
                        f"{label!r}: the {role!r} expression reads a "
                        f"variable whose own tracks have not "
                        f"materialized yet — attach that variable to "
                        f"the model BEFORE the one that reads it "
                        f"(adoption order is construction order)."
                    )
                ov = operand._value
                if ov is None:
                    raise ValueError(
                        f"{label!r}: the {role!r} expression reads a "
                        f"variable that has no value yet (still "
                        f"floating, or an unadopted mo.schedule) — "
                        f"attach it to the model first."
                    )
                if isinstance(ov, TrackValues):
                    if role not in ov:
                        raise ValueError(
                            f"{label!r}: the {role!r} expression reads "
                            f"a tracked variable without a {role!r} "
                            f"coordinate ({', '.join(ov.roles)}) — "
                            f"slice explicitly: .at(track='...')."
                        )
                    ov = ov[role]
                if self.track_expr(role) is not None:
                    # The track refers to that line (``actual=s.raw``): a
                    # formula, so a blank reads as 0 - kept blank in a
                    # line of dates or words.
                    from .regrain import hole_value
                    hole = hole_value(ov)
                    ov = ([hole if v is None else v for v in ov]
                          if isinstance(ov, list) else (hole if ov is None else ov))
                new_tracks[role] = (
                    list(ov) if isinstance(ov, list) else ov
                )
                # Identity membership — ``in`` would route through the
                # overloaded formula-building ``__eq__``.
                if not any(r is operand for r in self._dependency_refs):
                    self._dependency_refs.append(operand)
            else:
                new_tracks[role] = (
                    self._lower_role_cells(list(operand), role, label)
                    if isinstance(operand, list) else operand
                )
        if not new_tracks:
            return
        self._value = TrackValues(new_tracks)
        self._roles_materialized = True
        if self._value.time_length is not None:
            self.var_type = 'list'

    def _entered_mask(self, role, periods: int, shown: bool = False,
                      readings=None):
        """Per period, was track ``role`` - a formula - entered? None when
        its own blanks answer (a track entered by hand) or it reads no
        entries. A formula is a number even over months nobody entered (a
        blank counts as zero), so it is asked of the entries it is built
        from (``core.entries``). ``shown``: read through a lag as well — the mask
        a derived line keeps as ``_blend_entered``. ``readings``: the lines'
        readings shared across the model (``entries.readings_under``)."""
        from .tracks import TrackValues
        tv = self._value
        if (not isinstance(role, str) or not isinstance(tv, TrackValues)
                or role not in tv or not isinstance(tv[role], list)):
            return None
        if isinstance((self._role_kwargs or {}).get(role), list):
            # Entered as a list - of numbers, or a row of references read
            # cell by cell: its own blanks answer (``track_expr`` would
            # only look every cell over to say so).
            return None
        authored = self.track_expr(role)
        if authored is None and role in (self._role_kwargs or {}):
            return None
        expr = authored if authored is not None else self._expr
        if expr is None:
            return None
        from .entries import entered_mask, entry_lines, shown_entered_mask
        if shown:
            return shown_entered_mask(expr, role, periods, own=authored is not None,
                                      readings=readings)
        entries = entry_lines(expr, role, own=authored is not None, readings=readings)
        return entered_mask(entries, periods) if entries else None

    def _synthesize_blend(self, owner, window, decl=None) -> None:
        """Grow the blend's synthesized track on a DATA line.

        ``live[t] = given[t] if t ≤ boundary else follow[t]`` — with a
        per-CELL fallback to the other side when a side is missing or
        holds a hole (a side wins a cell only where it is present).
        Idempotent: re-running overwrites the synthesized track from the
        current sources. Pure formula results inherit live from operands
        (a formula evaluated over its operands' live tracks IS the live
        value) — except a DERIVED line with authored data, whose
        after-boundary tail continues the inherited live, not the follow
        track. Scalar and length-1 tracks splice as constants. Lines with
        their own start=/grain= are skipped (their positions are not the
        window's). ``decl``: the tracks declaration the caller resolved
        for ``owner`` already; resolved here when not given.
        """
        from .time import _parse, _period_labels
        from .tracks_decl import resolve_tracks_decl

        self._blend_entered = None
        tv = self._value
        if not isinstance(tv, TrackValues):
            return
        if getattr(self, '_regrained_tracks', False):
            # A coarser copy: each track, the blend among them, folded on
            # its own. Spliced again it would be no quarter of actuals, and
            # its blend is no authored track either.
            return
        label = (self._display_name
                 or getattr(self, '_python_name', None) or 'variable')
        if decl is None:
            decl = resolve_tracks_decl(owner)
        b = getattr(decl, 'blend', None) if decl is not None else None

        # A copy carried into a context that no longer synthesizes its
        # track (no declaration, no blend, or a renamed one) sheds the
        # stale series instead of detonating the mismatch law there.
        if self._blend_name is not None and (
                b is None or b.name != self._blend_name):
            if self._blend_name in tv:
                shed = tv.as_dict()
                shed.pop(self._blend_name, None)
                if shed:
                    self._value = tv = TrackValues(shed)
            self._blend_name = None
        if b is None:
            return

        is_data_line = (self._role_kwargs is not None
                        or self._expr is None)
        if not is_data_line:
            # A derived line computes its live value from its operands'
            # live values; its actual is still a formula that shows 0 over
            # months nobody entered; this mask records which were entered.
            given0 = tv[b.given] if b.given in tv else None
            if isinstance(given0, list):
                from .entries import readings_under
                self._blend_entered = self._entered_mask(
                    b.given, len(given0), shown=True,
                    readings=readings_under(owner))
            # A DERIVED line normally inherits live through the
            # broadcast — but one computed inside a CLASS BODY ran
            # before its operands had live (provisional
            # materialization). For a MEMORYLESS formula the splice of
            # its own per-track results IS the live value
            # (per-period: f(live_in[t]) = given-result[t] if t ≤ boundary
            # else follow-result[t]) — synthesize it. Stateful formulas
            # (recurrence/lag) are NOT splice-equal; they stay without
            # live and the mismatch law surfaces any cross-use loudly.
            tv0 = self._value
            if not isinstance(tv0, TrackValues):
                return
            # The declaration resolved above still holds: nothing since
            # has touched the containers.
            if b.name in tv0:
                return
            from .expr import SelfRef, FuncCall, MethodCall
            def _stateful(node) -> bool:
                stack = [node]
                seen0: set = set()
                while stack:
                    n = stack.pop()
                    if id(n) in seen0:
                        continue
                    seen0.add(id(n))
                    if isinstance(n, SelfRef):
                        return True
                    if isinstance(n, FuncCall) and getattr(
                            n, 'func', '').lower() in ('lag', 'lead'):
                        return True
                    if isinstance(n, MethodCall) and getattr(
                            n, 'method', '').lower() in (
                            'cumsum', 'shift', 'rolling_sum',
                            'rolling_mean', 'pct_change'):
                        return True
                    for attr in getattr(n, '__dataclass_fields__', {}):
                        v = getattr(n, attr, None)
                        if isinstance(v, (list, tuple)):
                            stack.extend(
                                x for x in v if hasattr(x, '__class__'))
                        elif hasattr(v, '__dataclass_fields__'):
                            stack.append(v)
                return False
            if _stateful(self._expr):
                return
            # fall through: splice this line's own tracks below
        if self._start is not None or self._grain is not None:
            return    # own-located line: window positions are not its
        # Authored-name collision, BOTH spellings: role kwargs / dotted
        # bindings are caught here via the stash; the tracks= dict
        # spelling lands as a role in the value with no synthesis flag.
        authored = (
            (self._role_kwargs is not None and b.name in self._role_kwargs)
            or (self._expr is None and b.name in tv
                and self._blend_name != b.name)
        )
        if authored:
            raise ValueError(
                f"{label!r} authors {b.name!r}, but the declaration "
                f"synthesizes that track (mo.blend name={b.name!r}) — "
                f"rename the authored track, or drop the blend."
            )
        if window.start is None or window.periods is None:
            return    # declaration-time validation demands a full window
        periods = window.periods
        if b.until is None:
            # Boundary-less splice: ``given`` wins wherever it carries a
            # value, ``follow`` fills the rest. Same rule the bounded
            # form applies BEFORE its boundary — with no date to keep
            # current, which suits a given track whose values arrive
            # irregularly and incompletely. No label is compared, so
            # none is spelled: the window counts its periods, and its
            # start is parsed as the labels would have parsed it, so a
            # malformed one fails the same way.
            _parse(window.start, window.grain)
            before_boundary = [True] * periods
        else:
            labels = _period_labels(
                window.start, None, periods, window.grain
            )
            try:
                until_anchor = _parse(b.until, window.grain)
            except (ValueError, TypeError) as exc:
                raise ValueError(
                    f"mo.blend until={b.until!r} is not a {window.grain} "
                    f"period label — the boundary must be spelled at the "
                    f"model grain (like the window start {window.start!r})."
                ) from exc
            # Compare PARSED anchors, never raw strings — '2025-2' would
            # sort after '2025-12' lexicographically and silently misroute
            # the whole splice.
            before_boundary = [
                _parse(lb, window.grain) <= until_anchor for lb in labels
            ]
        given = tv[b.given] if b.given in tv else None
        # A derived line that acquired authored data (dotted binding /
        # override spelling) continues its INHERITED live after the
        # boundary — the follow track would replay the output splice
        # this law exists to kill.
        prior_live = tv[b.name] if b.name in tv else None
        if self._expr is not None and prior_live is not None:
            follow = prior_live
        else:
            follow = tv[b.follow] if b.follow in tv else None
        if given is None and follow is None:
            return

        # A RECORD's scalar line (an amount, a date, a status) blends
        # to a SCALAR: given when present, else follow. The time
        # splice would spread the constant into a full-window series
        # and quietly turn a record field temporal — plan/actual/
        # forecast belong to record scalars too, without giving them a
        # timeline. OUTSIDE records a scalar-per-track line is an
        # ASSUMPTION whose live value switches at the boundary
        # (actual=100 → budget=90) — that splice stays. The container
        # marks itself: ``_records_container = True`` (any
        # records-shaped container may).
        def _in_record() -> bool:
            node = owner
            while node is not None:
                if getattr(node, '_records_container', False):
                    return True
                node = (getattr(node, '_parent', None)
                        or getattr(node, '_owner', None))
            return False
        if (not isinstance(given, list) and not isinstance(follow, list)
                and not any(isinstance(tr, list) for tr in tv.values())
                and _in_record()):
            merged = tv.as_dict()
            merged[b.name] = given if given is not None else follow
            self._value = TrackValues(merged)
            self._blend_name = b.name
            return

        def _cells(track) -> list:
            """One side, a cell per period: a missing side is holes
            throughout, a short list ends in holes."""
            if track is None:
                return [None] * periods
            if isinstance(track, list):
                if len(track) == 1:
                    return track * periods    # length-1 broadcasts, like a scalar
                if len(track) >= periods:
                    return track
                return track + [None] * (periods - len(track))
            return [track] * periods          # scalar splices as a constant

        from .entries import readings_under
        readings = readings_under(owner)

        def _entered(role, series):
            """Per period: was this side entered? (``_entered_mask``)"""
            if series is None:
                return None
            return self._entered_mask(role, periods, readings=readings)

        from .regrain import hole_value
        blank: list = []    # what neither side gives counts as, once a cell asks
        given_mask = _entered(b.given, given)
        follow_mask = (None if follow is prior_live
                       else _entered(b.follow, follow))
        # Per period, was the given side - a formula - entered? Kept as
        # ``_blend_entered``: the formula's number alone cannot tell an
        # entered zero from a month nobody entered.
        self._blend_entered = given_mask
        given_cells = _cells(given)
        follow_cells = _cells(follow)
        live = []
        for t in range(periods):
            if before_boundary[t]:
                first, second, mask = given_cells, follow_cells, given_mask
            else:
                first, second, mask = follow_cells, given_cells, follow_mask
            cell = first[t]
            if cell is None or (mask is not None and not mask[t]):
                cell = second[t]
            if cell is None:
                if not blank:
                    blank.append(hole_value(tv))
                cell = blank[0]         # neither side: the formula's 0, as in the book
            live.append(cell)
        merged = tv.as_dict()
        merged[b.name] = live
        # Length-1 lists broadcast (the extent law's "single value"
        # case); the synthesized full-window track would violate the
        # equal-length law beside them — materialize the broadcast.
        if periods > 1:
            for role, track in list(merged.items()):
                if isinstance(track, list) and len(track) == 1:
                    merged[role] = track * periods
        self._value = TrackValues(merged)
        self._blend_name = b.name
        if self.var_type != 'list':
            self.var_type = 'list'

    def __getattr__(self, name: str) -> 'Variable':
        """Dotted coordinate read — ``revenue.actual``.

        Fires only when normal attribute lookup MISSES, so every real
        Variable attribute and property wins automatically; a track
        that shadows one (a track literally named ``value``) stays
        reachable via ``.at(track=...)``, the canonical spelling this
        sugar delegates to.
        """
        if name.startswith('_'):
            raise AttributeError(name)
        from .tracks import TrackValues
        value = self.__dict__.get('_value')
        if isinstance(value, TrackValues):
            if name in value:
                return self.at(track=name)
            raise AttributeError(
                f"{self._display_name or 'Variable'!r} has no attribute "
                f"or track {name!r}; tracks: {', '.join(value.roles)}."
            )
        raise AttributeError(
            f"'Variable' object has no attribute {name!r}"
        )

    def track_expr(self, role: str) -> Expr | None:
        """The expression a track was AUTHORED with, or None.

        ``mo.Variable(plan=<Variable>)`` keeps the operand in the role
        stash while its VALUE lands in the track — the track's AST is
        the operand itself. A rooted operand renders as a reference to
        its cell; a floating expression inlines its own tree. Data
        tracks (lists, scalars, floating literals) and unauthored roles
        have no expression — they are inputs, and emit as values.
        """
        from .expr import ListExpr, Literal, VarRef
        operand = (self._role_kwargs or {}).get(role)
        if isinstance(operand, list):
            # A positional list with Variable ELEMENTS (a sparse row of
            # references: ``forecast=[actual_q1, None, …, other_cell, …]``) is a
            # compound formula — item i is period i's cell reference or
            # its literal. All-literal lists are data.
            if not any(isinstance(x, Variable) for x in operand):
                return None
            return ListExpr([
                VarRef(x) if isinstance(x, Variable) else Literal(x)
                for x in operand
            ])
        if not isinstance(operand, Variable):
            return None
        if operand._owner is None and (operand._expr is None or is_typed_date(operand)):
            return None                 # a typed value - a date typed with mo.DATE too
        return VarRef(operand)

    def shown_track_expr(self) -> Expr | None:
        """The expression behind the row a renderer SHOWS, or None.

        The shown track's own authored expression when it has one. When
        the shown track is the synthesized blend, the splice equals the
        FOLLOW track wherever the given track is absent — a line that
        authors only the follow (``forecast=<expr>`` with no actual) is
        that expression outright, so it renders as such. A splice with
        both sources authored is data (the blend writer's job), and a
        data track stays a value.
        """
        found = self.shown_track()
        return found[1] if found is not None else None

    def shown_track(self) -> "tuple[str, Expr] | None":
        """``(coordinate, expression)`` behind the shown row, or None.

        The coordinate is the track the expression was authored FOR —
        the follow track when the shown row is the blend standing in
        for it — which is what a renderer must resolve operands in.
        """
        shown = self.display_track_role()
        if shown is None:
            return None
        own = self.track_expr(shown)
        if own is not None:
            return (shown, own)
        from .tracks import TrackValues
        from .tracks_decl import resolve_tracks_decl
        decl = resolve_tracks_decl(self)
        blend = getattr(decl, 'blend', None)
        if blend is None or getattr(blend, 'name', None) != shown:
            return None
        given = getattr(blend, 'given', None)
        follow = getattr(blend, 'follow', None)
        if not isinstance(follow, str) or follow not in (self._role_kwargs or {}):
            return None
        # "Given absent" is a fact about the VALUE, not the kwargs stash:
        # a shared positional list (or a dotted binding) fills the given
        # track without stashing it, and that line is a real splice.
        value = self._value
        if (isinstance(value, TrackValues) and isinstance(given, str)
                and given in value):
            return None
        follow_expr = self.track_expr(follow)
        return (follow, follow_expr) if follow_expr is not None else None

    def display_track_role(self) -> str | None:
        """The role a renderer shows by default for this tracked line —
        the synthesized blend when the declaration blends, else the
        first authored role. None for an untracked Variable."""
        from .tracks import TrackValues
        from .tracks_decl import resolve_tracks_decl
        value = self._value
        if not isinstance(value, TrackValues):
            return None
        decl = resolve_tracks_decl(self)
        blend_name = getattr(getattr(decl, 'blend', None), 'name', None)
        roles = list(value.roles)
        return blend_name if blend_name in roles else roles[0]

    def _carry_presentation(self, result: 'Variable') -> 'Variable':
        """The line's unit, number format and note on its coarser copy, so a
        year sheet built from ``.at('year')`` shows the line's numbers as its
        month sheet does. The format is the line's own, worn only where the
        copy's sheet says none (``resolve_number_format``); the rest of the
        line's ``excel_props`` - hidden, a header row, an article number, a
        result's look - belongs to where the line stands, not to its copy."""
        if self._unit is not None and (result._unit is None
                                       or getattr(self, '_unit_declared', False)):
            result._unit = self._unit
            result._unit_declared = getattr(self, '_unit_declared', False)
        own = (self._excel_props or {}).get('number_format') or getattr(
            self, '_carried_number_format', None)
        if own:
            result._carried_number_format = own
        if self._description and not result._description:
            result._description = self._description
        return result

    def at(self, grain: Optional[str] = None, *,
           track: Optional[str] = None) -> 'Variable':
        """Polymorphic restrict/re-grain.

        ``at(grain)`` re-grains to a coarser grain by this Variable's
        own rule. ``at(track='actual')`` is the coordinate RESTRICT:
        drop the tracks axis, return the one track as an ordinary
        series (its own auditable AST node). Combined
        ``at('quarter', track='actual')`` restricts THEN re-grains.

        Time-agnostic Variables broadcast unchanged; ``grain == native``
        is a no-op copy. A located Variable needs a ``regrain=`` rule;
        grain-frozen series refuse.
        """
        self._refuse_pending()
        from .tracks import TrackValues
        label = self._display_name or self.id
        if track is not None:
            if (self._role_kwargs is not None
                    and not self._roles_materialized):
                raise ValueError(
                    f"{label!r} authors track coordinates that have not "
                    f"materialized yet — attach it to the model before "
                    f"slicing (.at(track=...) reads the adopted value)."
                )
            if not isinstance(self._value, TrackValues):
                raise ValueError(
                    f"{label!r} has no track coordinates — .at(track=...) "
                    f"restricts a Variable with tracks; this one is a plain "
                    f"series."
                )
            if track not in self._value:
                raise ValueError(
                    f"{label!r} has no coordinate {track!r}; declared: "
                    f"{', '.join(self._value.roles)}."
                )
            sliced = Variable(display_name=(
                f"{self._display_name} · {track}" if self._display_name
                else None
            ))
            sliced._set_expr(Restrict(VarRef(self), track))
            track = self._value[track]
            sliced._value = list(track) if isinstance(track, list) else track
            sliced.var_type = (
                'list' if isinstance(track, list) and len(track) > 1
                else 'scalar'
            )
            sliced._unit = self._unit
            sliced._regrain = self._regrain
            sliced._start = self._start
            sliced._grain = self._grain
            sliced._first_day = self._first_day
            # Keep the ambient chain reachable for .time/regrain on the
            # slice even before adoption.
            if getattr(self, '_owner', None) is not None:
                sliced.__dict__['_owner'] = self._owner
                # The slice runs on the container's window without being
                # one of its lines: a window set on the container later
                # moves it all the same (``multi_variable``).
                self._owner.__dict__['_sliced'] = True
            if grain is not None:
                return sliced.at(grain)
            return sliced
        if grain is None:
            raise TypeError(
                "at() needs a grain ('quarter'/'half'/'year') and/or a "
                "track= coordinate."
            )
        if isinstance(self._value, TrackValues):
            # Track-wise re-grain: each track projects by this
            # line's own rule (slice-then-regrain per role); the tracks
            # axis passes through unchanged.
            from .regrain import hole_value
            line_blank = hole_value(self._value)
            projected = {}
            p_start = None
            for role in self._value.roles:
                sliced = self.at(track=role)
                # What a hole counts as is the whole line's, not one
                # track's: actuals not entered yet in a line of dates.
                sliced._hole_value = line_blank
                p = sliced.at(grain)
                projected[role] = p._value
                if p_start is None:
                    p_start = p._start
            out = Variable(display_name=self._display_name)
            out._value = TrackValues(projected)
            # Each track - the blend among them - folded on its own: the
            # book must not splice the folded tracks again, a quarter of
            # actuals entered through February is no quarter of actuals.
            out._regrained_tracks = True
            out.var_type = (
                'list' if out._value.time_length is not None else 'scalar'
            )
            out._unit = self._unit
            out._grain = grain
            out._start = p_start
            out._first_day = getattr(p, '_first_day', None) if p is not None else None
            return self._carry_presentation(out)
        if self._indexed_by:
            # Rank-1 guard: without it, ``time`` resolving to None
            # would silently COPY the axised variable through the
            # projection — and downstream arithmetic with genuinely
            # re-grained series then misaligns coordinates against periods.
            raise ValueError(
                f"{label!r} is laid out along a finite axis — grain "
                f"projection of axised Variables is not supported. "
                f"Project a slice (drop the axis first), or keep "
                f"the model at native grain."
            )
        loc = self.time
        if loc is None:
            # A list-valued series that carries a re-grain rule but whose native
            # grain can't be resolved is a footgun: silently copying it leaves it
            # un-re-grained (a monthly series sitting in a "quarterly" view). Fail
            # loud instead — the rule has nothing to anchor to.
            if (isinstance(self._value, list) and self._regrain is not None
                    and not self._regrain.frozen):
                raise ValueError(
                    f"{label!r} has a re-grain rule but no resolvable native "
                    f"grain — set grain= on it, or default_grain on an ancestor "
                    f"model, so the rule knows which grain it re-grains from."
                )
            return self.copy()
        if loc.grain == grain:
            return self.copy()
        from .regrain import LINE_UNKNOWN
        spec = self._regrain
        values, start_label, first_day, ranges, rule = series_at_grain(
            label, self._value, loc, spec, grain,
            getattr(self, '_hole_value', LINE_UNKNOWN),
        )
        result = Variable(values, display_name=self._display_name,
                          start=start_label, grain=grain, regrain=spec)
        # A first bucket that starts inside its coarse period is short.
        result._first_day = first_day
        # Live formula: each target cell renders as a reducer over the source's
        # native cells (=SUM(range) / period-end cell), falling back to the
        # inlined value when the source has no address in the workbook.
        result._set_expr(Regrain(source=self, buckets=[(lo, hi) for _, lo, hi in ranges],
                                 recipe=rule, fill_values=list(values)))
        return self._carry_presentation(result)

    def set_regrain(self, spec: 'RegrainSpec') -> 'Variable':
        """Attach a re-grain rule to an operator-built Variable (chainable).

        ``revenue = (price * seats).set_regrain(mo.up('sum'))`` — formulas need
        a rule too, since re-grain applies each line's rule to its own value.
        """
        if not isinstance(spec, RegrainSpec):
            raise TypeError(
                "set_regrain takes a RegrainSpec — use mo.up(...), mo.frozen(), "
                "or mo.ratio(...)."
            )
        self._regrain = spec
        return self

    # Construction dispatchers (_resolve_keys, _init_from_formula,
    # _init_from_pyformula, _init_from_value, _build_compound_formula,
    # _coerce_literal, _auto_detect_types, _process_kwargs) live on
    # :class:`_VariableInit` in ``variable_init.py``. Variable inherits
    # them via the class declaration above.

    @property
    def path(self) -> QPath:
        """Qualified path identity — the single source of identity.

        Resolution order on an un-crystallized Variable:

        1. **Adopted as a named component** (``_owner`` + ``_component_name``
           set via ``_register_component``): derive from the owner's path
           and cache — so ``sheet.revenue.path`` is ``sheet.revenue``
           even when adoption happened before the sheet crystallized.
        2. **Top-level binding** (``mv = ...; mv.x = mo.Variable(...)``):
           the parent's path crystallizes when its own ``python_name``
           is set, and then case 1 applies.
        3. **Truly floating** (unbound, unadopted — e.g. an operator
           intermediate stashed in a list): return
           ``QPath(("__floating__", f"v{hex(id)}"))`` — honest about
           lacking structure.
        """
        if self._qualified_id is not None:
            return self._qualified_id
        owner = self._owner
        comp_name = self._component_name
        # Always derive from the owner when adopted — even if the
        # owner's path is currently floating. Otherwise two Variables
        # adopted under different floating owners (``a.x`` and ``b.x``)
        # both fall back to ``ROOT.x`` and collide. When the owner
        # later crystallizes, ``_crystallize_subtree`` rewrites the
        # cached _qualified_id to the rooted form.
        if owner is not None and comp_name:
            self._qualified_id = owner.path.child(comp_name)
            return self._qualified_id
        if self.python_name:
            self._qualified_id = QPath.ROOT.child(self.python_name)
            return self._qualified_id
        return QPath.floating(id(self), kind="v")

    # ``id`` property is inherited from :class:`Base`.

    @property
    def value(self) -> Any:
        """Current computed value, or ``None`` for backend-exclusive
        functions with no Python implementation.

        Returns ``self._value`` directly. Exposed as a property (not a
        bare attribute) so a subclass can load or evaluate the value on
        demand without touching callers. External code (user scripts,
        serializers, inspection tools) should use ``variable.value``.
        Internal engine code uses ``_value`` for directness.
        """
        self._refuse_pending()
        return self._value

    def _refuse_pending(self) -> None:
        """A value computed from a row that is not assigned yet is
        provisional — refuse to hand it out (see :mod:`modeleon.core.loops`)."""
        if self._awaits or self._forward_of is not None:
            from .loops import pending_message
            raise ValueError(pending_message(self))

    @property
    def formula(self) -> Optional[str]:
        """Formula string, derived from the expression tree.

        Rendered fresh on every access via ``_expr.to_string()``. Falls back
        to ``_raw_formula_str`` for formulas passed as bare strings, and to
        ``pyformula`` for pyformula-mode Variables. Returns None for plain
        input Variables. External consumers (JSON export, inspection
        tools) see the rendered string either way.
        """
        if self._expr is not None:
            return self._expr.to_string()
        if self._raw_formula_str is not None:
            return self._raw_formula_str
        if self.pyformula is not None:
            return self.pyformula
        return None

    @formula.setter
    def formula(self, value: Optional[str]) -> None:
        """Override the formula with a raw string.

        Used by consumers that need to post-process the rendered formula
        (e.g. a name-qualification pass that rewrites references). Setting
        a string drops the AST representation — the new string becomes
        the source of truth for ``.formula``, ``is_formula`` stays correct.
        """
        self._expr = None
        self._raw_formula_str = None if value is None else str(value)

    @property
    def source_code(self) -> Optional[str]:
        """Call-site rendering of the expression that produced this Variable.

        A sibling to :attr:`formula`, but showing a different view:

        - ``.formula`` — structural, rendered from the AST. ``VarRef``
          nodes print their referenced Variable's name. ``PMT(rate,
          nper, pv, ...)`` stays symbolic.
        - ``.source_code`` — captured at call time with resolved scalar
          values. The same ``PMT(...)`` call with rate=0.005, nper=360,
          pv=300000 yields ``"PMT(0.005, 360, 300000, 0, 0)"``.

        ``functions/`` helpers (financial, recurrence, cohort, …) set
        this for their results; plain inputs and operator-built
        intermediates leave it ``None``. The condition family (``IF``,
        ``AND``, ``OR``, ``NOT``, ``CHOOSE`` and the logical helpers)
        keeps no snapshot: its text is :attr:`formula` rendered when it is
        read, so it follows a later rename or a formula set by hand, as
        ``.formula`` does. Inspection tools read it via this property
        when they want the concrete call shape rather than the AST view.
        Returns ``None`` when no call-site snapshot was captured.
        """
        return self._source_code

    @source_code.setter
    def source_code(self, value: Optional[str]) -> None:
        """Override the call-site rendering — mostly used by extension
        authors building custom helpers. Setting to ``None`` clears it.
        """
        self._source_code = None if value is None else str(value)

    def _source_from_formula(self) -> None:
        # ``_source_code`` reads as the formula, rendered when read (see
        # ``_SourceCode``). A text written before (``elementwise`` stamps
        # the resolved call) would win over the rule, so it goes.
        self.__dict__.pop('_source_code', None)
        self.__dict__['_source_is_formula'] = True

    @property
    def is_formula(self) -> bool:
        """True when this Variable carries a formula (AST or raw string).

        ``pyformula`` is an input expressed via Python, not a formula, so it
        doesn't count here.
        """
        if self.pyformula is not None:
            return False
        return self._expr is not None or self._raw_formula_str is not None

    @property
    def expr(self) -> Optional[Expr]:
        """The expression tree (AST) that produces this Variable's value,
        or None for pure inputs. Assignment via ``variable.expr = node``
        automatically rebuilds ``_dependency_refs`` from the new tree.
        """
        return self._expr

    @expr.setter
    def expr(self, value: Optional[Expr]) -> None:
        """Attach an expression tree (or clear it) and refresh dependency
        refs in one step.

        Preserves order of first occurrence; deduplicates by object identity.
        Construction sites and operator methods assign via this setter so
        ``_expr`` and ``_dependency_refs`` stay consistent.
        """
        # Plain instance fields: no class declares a descriptor for any of
        # them, so they go straight into ``__dict__`` — what ``__setattr__``
        # does with them, minus a Python-level call per field. Every
        # operator result gets its expression through here.
        fields = self.__dict__
        fields['_expr'] = value
        if value is None:
            fields['_dependency_refs'] = []
            fields['_reads_time'] = False
            return
        seen = set()
        refs: List['Variable'] = []
        for var in value.iter_refs():
            key = id(var)
            if key not in seen:
                seen.add(key)
                refs.append(var)
        fields['_dependency_refs'] = refs
        # A re-grained value is computed FOR its own window — whatever time
        # its native source read lives in the source's window.
        reads_time = isinstance(value, TimeRef)
        if not reads_time and not isinstance(value, Regrain):
            for r in refs:
                if getattr(r, '_reads_time', False):
                    reads_time = True
                    break
        fields['_reads_time'] = reads_time
        # A formula over a row that is not assigned yet computes a
        # provisional value: remember what it waits on, and let the
        # placeholder find it when the loop closes.
        awaits = None
        for var in refs:
            if var._forward_of is not None:
                awaits = awaits or set()
                awaits.add(var)
            elif var._awaits:
                awaits = awaits or set()
                awaits.update(var._awaits)
        fields['_awaits'] = awaits
        if awaits:
            for placeholder in awaits:
                dependents = placeholder._forward_dependents
                if dependents is not None:
                    dependents.append(self)

    def _set_expr(self, expr: Expr) -> None:
        """Thin delegate to the :attr:`expr` setter. Kept for the many
        existing call sites (``x._set_expr(node)``) that would otherwise
        need a mechanical rewrite. New code should use ``x.expr = node``.
        """
        # ``expr`` is a property, which ``__setattr__`` hands straight to
        # ``object.__setattr__``; going there directly skips a Python-level
        # call on every operator.
        _object_setattr(self, 'expr', expr)

    def _as_operand(self, parenthesize: bool = False) -> Expr:
        """Return this Variable as an operand in a larger expression.

        Always produces a :class:`VarRef` so the translator can resolve it
        to the Variable's cell when one is assigned — ``x / y`` where
        ``x`` is a named formula Variable renders as ``=x_cell / y_cell``,
        not as ``=x_formula_inlined / y_cell``.

        When ``parenthesize`` is True (division, unary minus, …) AND the
        Variable carries its own formula tree, wrap the VarRef in a
        :class:`Paren` so precedence is preserved in the rendered string.
        For Variables that end up with a cell address, this adds a
        harmless pair of parens around the cell reference (``(C1) / C3``).
        For unnamed intermediates with no cell, the translator's VarRef
        fallback inlines the formula inside the parens — yielding
        ``(a + b) / c`` as expected.
        """
        if parenthesize and self._expr is not None:
            return Paren(VarRef(self))
        return VarRef(self)

    @property
    def dependencies(self) -> Set[str]:
        """Set of upstream variable paths — dotted-string form.

        Derived from ``_dependency_refs`` (the source of truth): each
        arithmetic operation / helper records the Variable objects it
        reads from, and this property projects them to their current
        ``.id`` strings on every access. Rooted Variables get qualified
        paths (``"acme.pnl.revenue"``); unrooted ones get their floating
        path (``"__floating__.v3"``).
        """
        return {ref.id for ref in self._dependency_refs}
    
    # ``python_name`` read-only property inherited from :class:`Base`.
    # ``excel_props`` property, ``set_style``, ``set_display_name``
    # inherited from :class:`Component`.

    # ═══════════════════════════════════════════════════════
    #  METADATA METHODS
    # ═══════════════════════════════════════════════════════

    @property
    def display_name(self) -> str:
        """Human-readable name, with fallback chain.

        - Explicit user-set label (``Variable(display_name="Revenue")`` or
          ``.set_display_name("Revenue")``) if present.
        - Otherwise the Python identifier humanized —
          ``total_revenue`` → ``"Total Revenue"``,
          ``XIRR`` → ``"XIRR"`` (all-caps preserved),
          ``Smth_smth2`` → ``"Smth Smth2"``.
        - Otherwise the auto-generated machine id humanized
          (``_var_47`` → ``"Var 47"``).

        Matches :class:`MultiVariableBase.display_name` so
        ``var.display_name`` and ``mv.display_name`` both give you "the
        name a human sees" without having to remember which field to
        prefer for which case.
        """
        if self._display_name:
            return self._display_name
        identifier = (
            self.python_name
            or self._component_name
            or self.path.leaf
        )
        return humanize_identifier(identifier)

    def set_type(self, var_type: str, value_type: Optional[str] = None) -> 'Variable':
        """
        Set the variable type

        Args:
            var_type: Variable type (``"scalar"`` or ``"list"``).
            value_type: Optional - data type (float, int, string)
        """
        self.var_type = var_type
        if value_type is not None:
            self.value_type = value_type
        return self
    
    def copy(self) -> 'Variable':
        """Create a copy of this Variable — emits a ``source.copy()`` formula.

        The copy's AST references the source via a ``MethodCall`` node, so
        the Excel output renders as the source cell (no actual value
        duplication in the spreadsheet). Value and metadata are carried
        forward so the copy behaves the same in Python.
        """
        result = Variable()
        result.expr = MethodCall(VarRef(self), "copy", [], {})
        result._value = self._value
        result.var_type = self.var_type
        result.value_type = self.value_type
        result._unit = self._unit
        result._unit_declared = getattr(self, '_unit_declared', False)
        result._keys = list(self._keys) if self._keys is not None else None
        result._keys_source = self._keys_source
        result._excel_props = dict(self._excel_props)
        result._indexed_by = self._indexed_by
        # Time location + re-grain rule must survive a copy (adoption clones via
        # copy, so losing these would silently strip a line's grain / rule).
        result._start = self._start
        result._grain = self._grain
        result._first_day = self._first_day
        result._regrain = self._regrain
        # Pending adoption-time state must survive too: an unmaterialized
        # schedule / extension / role stash would otherwise be silently
        # stripped by the clone-on-ownership-change path and leave a
        # valueless variable that never errors.
        result._schedule = (
            dict(self._schedule) if self._schedule is not None else None
        )
        result._extend = self._extend
        result._role_kwargs = (
            dict(self._role_kwargs) if self._role_kwargs is not None else None
        )
        result._roles_materialized = self._roles_materialized
        result._blend_name = self._blend_name
        return result

    def __copy__(self) -> 'Variable':
        """Support ``copy.copy(var)`` — delegates to :meth:`copy`.

        Added so ``from copy import copy; copy(var)`` works the same as
        ``var.copy()``. Useful when Variable flows through generic code
        that calls ``copy.copy`` on unknown objects.
        """
        return self.copy()
    
    # ═══════════════════════════════════════════════════════
    #  UNIT SUPPORT
    # ═══════════════════════════════════════════════════════
    
    @staticmethod
    def _resolve_unit(unit_arg):
        """Convert a unit argument (str, Unit, or None) to a Unit instance or None."""
        if unit_arg is None:
            return None
        from .unit import Unit
        if isinstance(unit_arg, Unit):
            return unit_arg
        if isinstance(unit_arg, str):
            return Unit(unit_arg)
        from .multi_variable import MultiVariableBase
        if isinstance(unit_arg, MultiVariableBase):
            # A container passed as the unit carries no unit: the row
            # keeps none (lenient, rather than a TypeError).
            return None
        raise TypeError(
            f"a unit is a string (a dimension: 'USD m') or a Unit "
            f"(mo.Unit.dimensionless('factor')); got {type(unit_arg).__name__}."
        )
    
    @property
    def unit(self):
        """The measurement unit of this variable, or None."""
        return self._unit
    
    @unit.setter
    def unit(self, value: "str | Unit | None") -> None:
        self._unit = self._resolve_unit(value)
        self._unit_declared = self._unit is not None

    def set_unit(self, unit: "str | Unit | None") -> "Self":
        """Set the measurement unit and return ``self`` for chaining —
        beside ``set_display_name`` / ``set_description`` / ``set_style``,
        so a line reads as one expression::

            fin.interest = (fin.opening * a.rate).set_unit('EUR').set_display_name('Interest')

        A string is a dimension; a word that labels a pure number is a
        ``mo.Unit.dimensionless('factor')``, passed as the object.
        """
        self.unit = unit
        return self

    @property
    def keys(self):
        """Named index labels for list values, or None."""
        return self._keys

    # Unit propagation (_unit_mul / _unit_div / _unit_add / _get_unit),
    # all arithmetic operators (+, -, *, /, **, %, //, unary +/-), and
    # all comparison operators (==, !=, <, <=, >, >=) live on
    # :class:`_VariableArithmetic` in ``variable_ops.py``. Variable
    # inherits them via the class declaration above.

    # ═══════════════════════════════════════════════════════
    #  SHIFT (lag / lead)
    # ═══════════════════════════════════════════════════════

    def shift(self, periods: int = 1, fill_value=0.0) -> 'Variable':
        """
        Shift values by specified number of periods (lag/lead).

        Works on list-type Variables (produced by mo.IF, recurrence, operators, etc.).

        Args:
            periods: Number of periods to shift (positive = lag, negative = lead)
            fill_value: Value to fill shifted positions (default: 0.0).
                        Can be a numeric value or a Variable object.

        Returns:
            New Variable with shifted values and formula tracking for Excel translation.

        Raises:
            ValueError: If called on a non-list Variable.
        """
        if not isinstance(self._value, list):
            raise ValueError(
                f"shift() only works on list Variables (time-series / ranges). "
                f"This Variable holds a scalar ({self._value!r}). "
                f"If you want to lag a scalar against itself, use "
                f"``mo.recurrence(start=x, formula='{{prev}}', periods=N)`` "
                f"to produce a list first."
            )

        original_fill_value = fill_value

        if isinstance(fill_value, Variable):
            from ..functions.lag import check_fill
            check_fill(fill_value, "shift(..., fill_value=")
            fill_value = fill_value._value

        if periods > 0:
            shifted = [fill_value] * periods + self._value[:-periods]
        elif periods < 0:
            n = -periods
            shifted = self._value[n:] + [fill_value] * n
        else:
            shifted = list(self._value)
        # A shifted blank reads as 0, as Excel's =previous_cell does -
        # kept blank only in a line of dates or words.
        from .regrain import hole_value
        blank = hole_value(self._value)
        shifted = [blank if v is None else v for v in shifted]

        # Build AST: shift(periods)  OR  shift(periods, fill_value=<...>)
        # Store the raw, typed fill value (already unwrapped from any
        # Variable above) so the renderer emits a correct Excel literal —
        # ``DATE(...)`` for dates, a bare number for numerics. Stringifying
        # here would lose the type and mis-render (quoted text / bare date).
        kwargs: Dict[str, Any] = {}
        if isinstance(original_fill_value, Variable):
            # A line's fill stays a reference to its cell, as mo.lag keeps
            # it: the workbook reads it live, and a loop container checks
            # its unit and settles the row by it when it closes.
            kwargs["fill_value"] = VarRef(original_fill_value)
        elif not (isinstance(original_fill_value, float) and original_fill_value == 0.0):
            kwargs["fill_value"] = fill_value

        result = Variable()
        result._set_expr(MethodCall(VarRef(self), "shift", [periods], kwargs))
        result._value = shifted
        result.var_type = 'list'
        result._keys = list(self._keys) if self._keys is not None else None
        # mo.lag's unit law, the one a loop container derives the row by
        # again: the source's unit, a fill in a unit is in the source's.
        from ..functions.lag import lag_unit
        try:
            unit = lag_unit(self, original_fill_value, result)
        except ValueError as exc:
            raise ValueError(str(exc).replace("mo.lag takes", "shift() takes", 1)
                             .replace("in mo.lag,", "in .shift(),", 1)) from None
        if unit is not None:
            result._unit = unit

        return result

    # ═══════════════════════════════════════════════════════
    #  SPECIAL METHODS
    # ═══════════════════════════════════════════════════════
    
    def __index__(self) -> int:
        """Allow Variable to be used where Python expects an integer (range(), slicing, etc.).
        Only works for scalar Variables with integer-compatible values."""
        self._refuse_pending()
        v = self._value
        if isinstance(v, list):
            if len(v) == 1:
                v = v[0]
            else:
                raise TypeError(
                    f"Cannot convert a list Variable (length {len(v)}) to int. "
                    f"Pick a position first: ``int(var[0])`` or ``int(var[-1])``, "
                    f"or use ``var.value`` to get the list directly."
                )
        return int(v)

    def __int__(self) -> int:
        self._refuse_pending()
        v = self._value
        if isinstance(v, list):
            v = v[0] if len(v) == 1 else v
        return int(v)

    def __float__(self) -> float:
        self._refuse_pending()
        v = self._value
        if isinstance(v, list):
            v = v[0] if len(v) == 1 else v
        return float(v)

    def __bool__(self) -> bool:
        """Truthiness of the computed value — Python-mode branching works.

        ``if seats > 300:`` branches on the comparison's eager value
        (comparisons compute at operator time), so plain
        Python control flow behaves exactly as Python users expect.
        Note the mode difference: a Python ``if`` decides NOW, in
        Python — the condition does not travel into the emitted Excel
        workbook. Use ``mo.IF(cond, then, otherwise)`` when the
        condition should live in the model (and in Excel) as a formula.

        Without it, ``bool()`` would fall back to ``__len__`` (1 for
        scalars) — every ``if var:`` branch would run regardless of the
        value, silently.

        List-valued Variables refuse, pandas/numpy-style: the truth of
        many values at once is genuinely ambiguous.
        """
        self._refuse_pending()
        if isinstance(self._value, list):
            label = self.display_name or "Variable"
            raise ValueError(
                f"The truth value of a list Variable ({label!r}) is "
                "ambiguous. Reduce it first (e.g. mo.SUM(...), var[0], "
                "or a comparison on a single period)."
            )
        return bool(self._value)

    def __len__(self) -> int:
        """
        Get the length of a Variable (for list/array variables)

        Returns:
            Length of the value if it's a list, otherwise 1

        Example:
            years = Variable(display_name="Years", value=[2025, 2026, 2027, 2028, 2029], var_type='list')
            num_years = len(years)  # Returns 5

            scalar_var = Variable(display_name="Revenue", value=1000000)
            length = len(scalar_var)  # Returns 1
        """
        if isinstance(self._value, list):
            return len(self._value)
        return 1
    
    def __getitem__(self, key):
        """Subscript: positional index, slice, or named key.

        Dispatches to ``_get_by_key`` when ``self`` is keyed and ``key``
        names an entry, else to ``_get_by_index`` (scalar or list position)
        or ``_get_slice`` (range of positions on a list).

        Example::

            years = Variable([2025, 2026, 2027, 2028, 2029])
            first_year = years[0]         # scalar Variable, value=2025
            mid_years  = years[1:4]       # list Variable,   value=[2026, 2027, 2028]

            scenarios = Variable({"Bear": 800, "Base": 1000, "Bull": 1200})
            base = scenarios["Base"]      # scalar Variable, value=1000
        """
        # Keyed access: str keys always dispatch here; int keys only when
        # they look like labels (value out of positional range), so
        # revenue[2025] works for a keyed-by-year Variable.
        if self._keys is not None and not isinstance(key, slice):
            keyed = (
                isinstance(key, str)
                or (not isinstance(key, int) and key in self._keys)
                or (isinstance(key, int) and key in self._keys and key >= len(self._value))
            )
            if keyed:
                return self._get_by_key(key)

        if isinstance(key, slice):
            return self._get_slice(key)
        return self._get_by_index(key)

    def _get_by_key(self, key):
        """Named-key lookup: returns a scalar Variable pointing at the match."""
        # Caller (``__getitem__``) only invokes this branch when ``self._keys``
        # is truthy. The assert pins that contract for type-checkers.
        assert self._keys is not None
        if key not in self._keys:
            raise KeyError(
                f"Key {key!r} not found on this Variable. "
                f"Available keys: {self._keys}. "
                f"(This Variable was built from a dict or with ``keys=...``; "
                f"use one of the labels above, or a positional index.)"
            )
        idx = self._keys.index(key)
        result = Variable()
        result._set_expr(Subscript(VarRef(self), key))
        result._value = self._value[idx]
        result.var_type = 'scalar'
        result.value_type = self.value_type
        result._unit = self._unit
        return result

    def _get_by_index(self, key):
        """Positional index: returns a scalar Variable.

        For a scalar source, only ``[0]`` is valid (a convenience so
        arithmetic producing a scalar still looks list-ish). Non-zero
        indices on a scalar raise ``IndexError``. A tracked row's period
        is that period of every track — one live cell per track.
        """
        if isinstance(self._value, TrackValues):
            # ``closing[-1]`` of a plan/actual row: each track's own last
            # period (a scalar track is the same in every period), written
            # as one reference per track. It used to raise (an index on a
            # scalar) or, for ``[0]``, copy the whole row as numbers.
            from .regrain import hole_value
            blank = hole_value(self._value)
            picked: Dict[str, Any] = {}
            for role, track in self._value.items():
                cell = track[key] if isinstance(track, list) else track
                if cell is None:
                    # A period not entered yet (an actual still to come)
                    # reads as a blank cell does in a formula: 0 — but a
                    # line of dates or words has no zero.
                    if blank is None:
                        raise IndexError(
                            f"Period {key} of {self.id!r} is blank in its "
                            f"{role!r} track, and a line of dates or words "
                            f"has no zero to stand in for it."
                        )
                    cell = blank
                picked[role] = cell
            result = Variable()
            result._set_expr(Subscript(VarRef(self), key))
            result._value = TrackValues(picked)
            result.var_type = 'scalar'
            result.value_type = self.value_type
            result._unit = self._unit
            return result
        if not isinstance(self._value, list):
            if key == 0:
                return Variable(value=self._value, var_type=self.var_type, value_type=self.value_type)
            raise IndexError(
                f"Index {key} on a scalar Variable — only ``[0]`` is valid "
                f"(as a convenience for list-shaped APIs). A scalar row is "
                f"one cell: read it as it is, by its name."
            )

        result = Variable()
        result._set_expr(Subscript(VarRef(self), key))
        result._value = self._value[key]
        result.var_type = 'scalar'
        result.value_type = self.value_type
        result._unit = self._unit
        return result

    def _get_slice(self, key: slice):
        """Slice a list Variable: returns a list Variable with the range."""
        result = Variable()
        result._set_expr(Subscript(VarRef(self), key))
        result._value = self._value[key]
        result.var_type = 'list'
        result.value_type = self.value_type
        result._unit = self._unit
        if self._keys is not None:
            result._keys = self._keys[key]
        return result
    
    def __iter__(self):
        """
        Make Variable iterable - allows using it in for loops and list comprehensions
        
        Returns:
            Iterator over the values
        
        Example:
            revenue = Variable(display_name="Revenue", value=[1000, 1100, 1200], var_type='list')
            for r in revenue:  # Iterates over [1000, 1100, 1200]
                print(r)
            
            # Works in list comprehensions
            doubled = [r * 2 for r in revenue]  # [2000, 2200, 2400]
            
            # Scalar variables iterate once (yield their single value)
            scalar = Variable(display_name="Scalar", value=100)
            for s in scalar:  # Iterates once, yields 100
                print(s)
        """
        self._refuse_pending()
        if isinstance(self._value, list):
            return iter(self._value)
        else:
            # Scalar: return iterator that yields the single value once
            return iter([self._value])
    
    def __repr__(self):
        return f"Variable({self.id}={self._format_value_for_repr()})"

    def _format_value_for_repr(self) -> str:
        """Compact value rendering for ``__repr__`` / text display.

        Keeps the repr short and scannable. Long lists truncate to first
        three / last one with an ellipsis.
        """
        v = self._value
        if self._awaits or self._forward_of is not None:
            return "<pending>"
        if v is None:
            return "None"
        if isinstance(v, list):
            if len(v) <= 6:
                return "[" + ", ".join(self._format_scalar(x) for x in v) + "]"
            head = ", ".join(self._format_scalar(x) for x in v[:3])
            tail = self._format_scalar(v[-1])
            return f"[{head}, ..., {tail}]"
        return self._format_scalar(v)

    @staticmethod
    def _format_scalar(v) -> str:
        if v is None:
            return "None"
        if isinstance(v, bool):
            return str(v)
        if isinstance(v, int):
            return f"{v:,}"
        if isinstance(v, float):
            if v != v:  # NaN
                return "nan"
            if abs(v) >= 1000:
                return f"{v:,.0f}"
            return f"{v:g}"
        return repr(v)

    def _repr_html_(self) -> str:
        """Jupyter / IPython rich-display: render as a compact HTML table."""
        from ..display.html import variable_html
        return variable_html(self)

    def to_excel(self, path: "str | Path") -> None:
        """Emit this Variable to an ``.xlsx`` file as a one-row sheet.

        Mirrors ``MultiVariable.to_excel`` so a lone Variable can be
        exported without first wrapping it in an MV. The Variable lays
        out under a tab named after its display_name (humanized
        ``python_name`` when adopted, literal ``"Variable"`` when
        anonymous) — the same naming the HTML repr uses, so the file
        and the notebook view agree.
        """
        from ..compile.excel.layout import LayoutEngine
        from ..compile.excel.writer import to_excel as _to_excel
        _to_excel(path, root=LayoutEngine._make_variable_sheet(self))

