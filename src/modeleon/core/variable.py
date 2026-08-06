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
from typing import Optional, Set, Any, Dict, List, Sequence, Tuple, TYPE_CHECKING

from .humanize import humanize_identifier
from .component import Component
from .expr import Expr, MethodCall, Paren, Regrain, Restrict, Subscript, VarRef
from .qpath import QPath
from .regrain import RegrainSpec
from .shape import Shape
from .variable_init import _VariableInit
from .variable_ops import _VariableArithmetic

if TYPE_CHECKING:
    from pathlib import Path

    from .multi_variable import MultiVariableBase


#: A Variable's native time location — ``(start, grain)`` — or ``None`` when
#: time-agnostic. Returned by :attr:`Variable.time`.
TimeLoc = namedtuple('TimeLoc', ['start', 'grain'])


class Variable(_VariableInit, _VariableArithmetic, Component):
    # ``__eq__`` is overridden on ``_VariableArithmetic`` to return a
    # comparison Variable (formula building), not a bool. Python drops
    # the default ``__hash__`` when ``__eq__`` is overridden — restore
    # identity-based hashing so Variables can be members of sets,
    # WeakSets, and dict keys.
    __hash__ = object.__hash__

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
        # Data formatting
        'number_format',
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
                meaningful at the MV layer for sheet/template
                decisions). See ``modeleon_pro.excel_layout`` for the
                dataclass shape; engine stays duck-typed.
            indexed_by: Sequence of axis Variables this Variable is
                laid out along. An axis is just another Variable —
                typically one with a list of labels. Operator-built
                Variables get this set automatically by broadcast; pure
                inputs declare it here. Default ``()`` means scalar.
            **kwargs: Plugin-registered kwargs (e.g. ``control=`` from a
                pro extension). Unknown kwargs raise TypeError so typos
                don't silently disappear.

        ``value``, ``formula``, and ``pyformula`` are mutually exclusive.
        For arithmetic-produced Variables (``total = a + b``) the formula is
        built by the operator methods and ``__init__`` is called with
        ``formula=<string>``.
        """
        # Dict sugar: Variable({"Bear": 800, "Base": 1000}) → list+keys
        if isinstance(value, dict):
            keys = list(value.keys())
            value = list(value.values())

        if sum(arg is not None for arg in (value, formula, pyformula)) > 1:
            raise ValueError(
                "Cannot specify more than one of 'value', 'formula', or 'pyformula'. "
                "Use 'value' for literal values, 'formula' for calculations."
            )

        # --- track-kwarg authoring (§16.2/§16.10 — P2) ---
        # ``mo.Variable(факт=[...], бюджет=<expr>)`` — the coordinate
        # kwargs. Track names are USER CONTENT (declared once on the
        # model via ``tracks=mo.Tracks(...)``), so they cannot be
        # literal parameters; the stash is filled from the leftover
        # **kwargs after plugin consumption (see the ``_process_kwargs``
        # call below) and validated against the ambient declaration at
        # ADOPTION — the constructor only records intent.
        # The pin spelling (§16.10): a positional shared expression plus
        # track overrides — ``mo.Variable(price * seats, факт=ledger)``
        # — derives non-overridden tracks from the shared value.
        self._role_kwargs: Optional[dict] = None
        # True once adoption has lowered the track kwargs onto a
        # TrackValues. The guard must be a FLAG, not a check on the
        # value's type — a tracked SHARED expression makes ``_value``
        # TrackValues at construction, and the overrides still owe a
        # materialization pass.
        self._roles_materialized: bool = False
        # True while the value was materialized from the KWARGS' OWN
        # names, before any declaration was reachable (inside a
        # MultiVariableClass body, where the instance is still
        # floating). Adoption validates against the declaration and
        # clears the flag.
        self._roles_provisional: bool = False
        # Name of the track the BLEND synthesized on this variable (None
        # until synthesis). Distinguishes synthesized-by-us from
        # authored/inherited on re-runs, and lets a copy carried into a
        # blend-less context strip its now-stale synthesized track.
        self._blend_name: Optional[str] = None

        # --- tracked construction (§14.2.2 / §16 — the tracks value) ---
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
                    "tracks= IS the finite-axis value — indexed_by= is "
                    "reserved for future non-tracks axes and cannot be "
                    "combined with it (axis budget is one, §16.1)."
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
        super().__init__(
            display_name=display_name,
            excel_props=excel_props,
            description=description,
        )
        # Opaque layout-config attribute. Writer reads ``.format`` off it
        # to apply per-cell number_format overrides. ``None`` is the
        # inherit-from-parent signal in pro's resolution walker.
        self._excel_layout = excel_layout
        # Adoption back-pointers: set by ``MultiVariableBase._register_component``
        # when this Variable is attached to a parent MV. Default ``None`` so
        # the ``.path`` property can read them without ``getattr`` fallbacks.
        self._owner: Optional["MultiVariableBase"] = None
        self._component_name: Optional[str] = None
        self._value: Any = None
        # _expr is the structured source of truth for this Variable's
        # formula. The ``.formula`` string property is a read-through
        # view derived from ``_expr.to_string()`` — set ``_expr``
        # directly, not ``formula``.
        self._expr: Optional[Expr] = None
        # Fallback string form for formulas passed or assigned as raw
        # strings — e.g. ``Variable(formula="x + y")`` or external tools
        # writing into ``self.formula = <rendered>`` via the setter.
        # Internal DSL construction builds an AST on ``_expr`` and leaves
        # this None; the ``formula`` getter reads ``_expr`` first and
        # falls back here.
        self._raw_formula_str: Optional[str] = None
        # Captured rendering of a ``pyformula=`` input (see kwarg
        # docstring). Marks the Variable as an input, not a computed
        # formula. None for every other construction path.
        self.pyformula: Optional[str] = None
        # Human-readable source snippet for the expression that produced
        # this Variable (e.g. ``"cohort_retention(new_cust, rate)"``,
        # ``"SUM(revenues)"``, ``"recurrence(100, ...)"``). Set by
        # ``functions/*`` helpers so inspection tools can show a
        # friendlier rendering than the raw formula AST. Left as None
        # for Variables that don't need one — operator-overloaded
        # intermediates and plain inputs.
        self._source_code: Optional[str] = None
        # _dependency_refs mirrors _expr.iter_refs() (with dedup); populated by
        # construction sites that build an AST. Kept as a cached list rather
        # than a live property so callers don't re-walk the tree on every read.
        self._dependency_refs: List['Variable'] = []
        self._unit = self._resolve_unit(unit)
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
        self.var_type = var_type
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
        self._indexed_by: Tuple['Variable', ...] = tuple(indexed_by) if indexed_by else ()

        # --- axised-construction invariant (multi-axis §14.4) ---
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

        # --- ambient time location + re-grain (multi-axis-engine-design.md §5.1) ---
        # A Variable is DEFINED at a native (start, grain) — re-grainable
        # metadata — or is time-agnostic (no grain; broadcasts unchanged).
        self._start = start
        self._grain = grain
        self._regrain = regrain
        # ``mo.schedule`` steps ({date: value}, materialized at adoption
        # when the ambient window resolves) and the declared out-of-extent
        # behavior (mo.zero()/mo.hold()/mo.none(), §16.5) — zero/hold
        # materialize a partial series to the window at adoption; full
        # date-aligned arithmetic lands with the track layer.
        from .extend import ExtendRule
        self._schedule: Optional[dict] = None
        if extend is not None and not isinstance(extend, ExtendRule):
            raise TypeError(
                "extend= takes mo.zero(), mo.hold() or mo.none() — the "
                "declared continuation of a partial series beyond its "
                "values."
            )
        self._extend: Optional[str] = extend.kind if extend is not None else None
        # A positional Variable carries its PENDING materialization
        # state into the copy: ``mo.Variable(mo.schedule({...}),
        # actual=...)`` binds the schedule Variable before it ever adopts
        # (its ``_value`` is still None), so the wrapper inherits the
        # steps / extension rule and materializes them itself.
        # (``self._schedule`` was defaulted above, after the value
        # dispatch — the carry must live here, not in
        # ``_init_from_value``, or the default would clobber it.)
        if isinstance(value, Variable):
            if getattr(value, '_schedule', None) is not None:
                self._schedule = dict(value._schedule)
            if self._extend is None:
                self._extend = getattr(value, '_extend', None)
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
                    "mo.Variable(<shared formula>, факт=[...])."
                )
            if indexed_by:
                raise ValueError(
                    "track kwargs ARE the finite axis — indexed_by= "
                    "cannot be combined with them (axis budget is one, "
                    "§16.1)."
                )
            self._role_kwargs = _track_candidates
            # Materialize PROVISIONALLY from the kwargs' own names.
            # The declaration lives on the model, which a Variable
            # built inside a MultiVariableClass body cannot reach —
            # and the very next line (``self.выручка = self.часы *
            # ставка``) computes eagerly (ADR-010). The kwargs already
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

        AXISED Variables have no single time location (§17 rank-1 guard):
        the flat value enumerates COORDINATES, not periods — claiming an
        inherited monthly window over it is how the grain lens summed
        scenario cells into quarters with no error. Tracks (P1) give each
        coordinate its own series; until then the variable is time-less.
        """
        if self._indexed_by:
            return None
        if self._grain is not None:
            return TimeLoc(self._start, self._grain)
        # Only list-valued (time-series) Variables inherit an ancestor's
        # default grain; a scalar constant is time-agnostic — it broadcasts
        # unchanged into any projection — regardless of a model default.
        # A FIBERED value with list tracks is a time series too — time
        # lives INSIDE each track (the two-tier spine, §16.1).
        from .tracks import TrackValues
        if isinstance(self._value, TrackValues):
            if self._value.time_length is None:
                return None
        elif not isinstance(self._value, list):
            return None
        seen: Set[int] = set()
        node = getattr(self, '_owner', None) or getattr(self, '_parent', None)
        while node is not None and id(node) not in seen:
            seen.add(id(node))
            grain = getattr(node, 'default_grain', None)
            if grain is not None:
                return TimeLoc(getattr(node, 'default_start', None), grain)
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
        2. **The extent law (§16.6)**: a bare list must be length 1 or
           exactly the window length. A partial series is legal only when
           declared partial (own ``start=``, a schedule, or ``extend=``).
           Verified live: a 3-element list in a 24-period model was
           accepted silently and detonated at an unrelated line.

        Idempotent; a no-op while the window is unresolvable (floating
        subtrees get their check at crystallization).
        """
        from .time import resolve_default_window, _period_labels
        owner = getattr(self, '_owner', None) or getattr(self, '_parent', None)
        if owner is None:
            return
        window = resolve_default_window(owner)
        # Scalar track coordinates need no window (time lives INSIDE a
        # track; an all-scalar TrackValues is time-agnostic), so the
        # window gate scopes the TIME duties below — it must not swallow
        # role materialization on a windowless model.
        have_window = window is not None and window.grain is not None

        from .tracks import TrackValues
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
            # No return: a schedule can be the SHARED value under role
            # kwargs (pin spelling) — materialization continues below.

        # --- declared extension: materialize a partial series ---
        # ``extend=mo.zero()/mo.hold()`` legalizes the prefix spelling:
        # the author SAID what lies beyond the values, so the series
        # completes to the window here — the same moment a schedule
        # materializes. ``none`` stays un-materialized (its meaning is
        # absent/NA, which needs the track layer's semantics).
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

        # --- role materialization (P2, §16.2/§16.10) ---
        # ``mo.Variable(факт=..., бюджет=...)`` recorded intent at
        # construction; the ambient ``mo.Tracks`` declaration is
        # reachable only from the tree, so the lowering happens here.
        # Idempotence rides the ``_roles_materialized`` FLAG — a tracked
        # shared expression makes ``_value`` TrackValues at construction
        # already, and its overrides still owe this pass.
        # A provisionally-materialized variable still owes this pass:
        # its names came from the kwargs, and only the declaration can
        # say whether they are TRACKS or typos.
        if self._role_kwargs is not None and (
                not self._roles_materialized or self._roles_provisional):
            from .tracks_decl import resolve_tracks_decl
            label = (self._display_name
                     or getattr(self, '_python_name', None) or 'variable')
            decl = resolve_tracks_decl(owner)
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
                    f"window. (Ragged actuals arrive with the role-window "
                    f"law, P2.)"
                )
            # --- live synthesis (§16.9) — AFTER the extent law, so a
            # ragged track dies with the teaching error, never a raw
            # IndexError from the splice loop.
            self._synthesize_blend(owner, window)
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
    # descriptor is a TRACK BINDING (§16.10 MV-style write).
    _PLAIN_ATTRS = frozenset({'var_type', 'value_type', 'pyformula'})

    def __setattr__(self, name: str, value: Any) -> None:
        """Dotted track write — ``выручка.факт = [...]`` (§16.10).

        The MV-style incremental spelling: the fact arrives AFTER the
        line is defined, as its own statement, without rewriting the
        definition. Internal attributes (underscore), engine-assigned
        plain attributes, and class descriptors (properties) take the
        normal path; any other name is a track binding, validated
        against the ambient declaration when the tree can reach it and
        stashed until adoption when it cannot.
        """
        if (name.startswith('_') or name in self._PLAIN_ATTRS
                or hasattr(type(self), name)):
            super().__setattr__(name, value)
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
            v = list(v)
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
            # A flat shared value expands into every declared track
            # (§16.10 pin semantics), then the binding overrides one.
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
            self._synthesize_blend(owner, window)

    def _validate_tracks_against(self, decl, label: str) -> None:
        """Every authored track name must be declared (§16.2)."""
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
                # The pin spelling (§16.10): non-overridden roles
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
                new_tracks[role] = (
                    list(ov) if isinstance(ov, list) else ov
                )
                # Identity membership — ``in`` would route through the
                # overloaded formula-building ``__eq__``.
                if not any(r is operand for r in self._dependency_refs):
                    self._dependency_refs.append(operand)
            else:
                new_tracks[role] = (
                    list(operand) if isinstance(operand, list)
                    else operand
                )
        if not new_tracks:
            return
        self._value = TrackValues(new_tracks)
        self._roles_materialized = True
        if self._value.time_length is not None:
            self.var_type = 'list'

    def _synthesize_blend(self, owner, window) -> None:
        """Grow the blend's synthesized track on a DATA line (§16.9).

        ``live[t] = given[t] if t ≤ boundary else follow[t]`` — with a
        per-CELL fallback to the other side when a side is missing or
        holds a hole (§16.9's "and present" clause). Idempotent:
        re-running overwrites the synthesized track from the current
        sources. Pure formula results inherit live from operands (that
        inheritance IS the live-universe evaluation) — except a DERIVED
        line with pinned data, whose after-boundary tail continues the
        inherited live, not the follow track. Scalar and length-1
        tracks splice as constants. Lines with their own start=/grain=
        are skipped in v1 (their positions are not the window's).
        """
        from .time import _parse, _period_labels
        from .tracks import TrackValues
        from .tracks_decl import resolve_tracks_decl

        tv = self._value
        if not isinstance(tv, TrackValues):
            return
        label = (self._display_name
                 or getattr(self, '_python_name', None) or 'variable')
        decl = resolve_tracks_decl(owner)
        b = getattr(decl, 'blend', None) if decl is not None else None

        # A copy carried into a context that no longer synthesizes its
        # track (no declaration, no blend, or a renamed one) sheds the
        # stale series instead of detonating the mismatch law there.
        if self._blend_name is not None and (
                b is None or b.name != self._blend_name):
            if self._blend_name in tv.roles:
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
            # A DERIVED line normally inherits live through the
            # broadcast — but one computed inside a CLASS BODY ran
            # before its operands had live (provisional
            # materialization). For a MEMORYLESS formula the splice of
            # its own per-track results IS the live-universe value
            # (per-period: f(live_in[t]) = факт-result[t] if t ≤ close
            # else план-result[t]) — synthesize it. Stateful formulas
            # (recurrence/lag) are NOT splice-equal; they stay without
            # live and the mismatch law surfaces any cross-use loudly.
            tv0 = self._value
            from .tracks import TrackValues as _TV
            if not isinstance(tv0, _TV):
                return
            decl0 = resolve_tracks_decl(owner) if True else None
            b0 = getattr(decl0, 'blend', None) if decl0 is not None else None
            if b0 is None or b0.name in tv0.roles:
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
                            n, 'name', '').lower() in ('lag', 'lead'):
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
            or (self._expr is None and b.name in tv.roles
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
        try:
            until_anchor = _parse(b.until, window.grain)
        except (ValueError, TypeError) as exc:
            raise ValueError(
                f"mo.blend until={b.until!r} is not a {window.grain} "
                f"period label — the boundary must be spelled at the "
                f"model grain (like the window start {window.start!r})."
            ) from exc
        labels = _period_labels(
            window.start, None, window.periods, window.grain
        )
        # Compare PARSED anchors, never raw strings — '2025-2' would
        # sort after '2025-12' lexicographically and silently misroute
        # the whole splice.
        before_boundary = [
            _parse(lb, window.grain) <= until_anchor for lb in labels
        ]
        given = tv[b.given] if b.given in tv else None
        # A derived line that acquired pinned data (dotted binding /
        # pin spelling) continues its INHERITED live after the
        # boundary — the follow track would replay the output splice
        # this law exists to kill.
        prior_live = tv[b.name] if b.name in tv.roles else None
        if self._expr is not None and prior_live is not None:
            follow = prior_live
        else:
            follow = tv[b.follow] if b.follow in tv else None
        if given is None and follow is None:
            return

        def _cell(track, t):
            if track is None:
                return None
            if isinstance(track, list):
                if len(track) == 1:
                    return track[0]     # length-1 broadcasts, like a scalar
                return track[t] if t < len(track) else None
            return track                # scalar splices as a constant

        live = []
        for t in range(len(labels)):
            first, second = ((given, follow) if before_boundary[t]
                             else (follow, given))
            cell = _cell(first, t)
            if cell is None:
                cell = _cell(second, t)
            live.append(cell)
        merged = tv.as_dict()
        merged[b.name] = live
        # Length-1 lists broadcast (the extent law's "single value"
        # case); the synthesized full-window track would violate the
        # equal-length law beside them — materialize the broadcast.
        if window.periods > 1:
            for role, track in list(merged.items()):
                if isinstance(track, list) and len(track) == 1:
                    merged[role] = track * window.periods
        self._value = TrackValues(merged)
        self._blend_name = b.name
        if self.var_type != 'list':
            self.var_type = 'list'

    def __getattr__(self, name: str) -> 'Variable':
        """Dotted coordinate read — ``выручка.факт`` (§0's promise).

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

    def at(self, grain: Optional[str] = None, *,
           track: Optional[str] = None) -> 'Variable':
        """Polymorphic restrict/re-grain (§14.2.9 / §17.2).

        ``at(grain)`` keeps its shipped meaning — re-grain to a coarser
        grain by this Variable's own rule. ``at(track='actual')`` is the
        coordinate RESTRICT: drop the tracks axis, return the one track
        as an ordinary series (its own auditable AST node). Combined
        ``at('quarter', track='actual')`` restricts THEN re-grains.

        Time-agnostic Variables broadcast unchanged; ``grain == native``
        is a no-op copy. A located Variable needs a ``regrain=`` rule;
        grain-frozen series refuse.
        """
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
            # Keep the ambient chain reachable for .time/regrain on the
            # slice even before adoption.
            if getattr(self, '_owner', None) is not None:
                sliced.__dict__['_owner'] = self._owner
            if grain is not None:
                return sliced.at(grain)
            return sliced
        if grain is None:
            raise TypeError(
                "at() needs a grain ('quarter'/'year') and/or a "
                "track= coordinate."
            )
        from .regrain import Ratio
        from .time import regrain_series
        if isinstance(self._value, TrackValues):
            # P2.5 — track-wise re-grain: each track projects by this
            # line's own rule (slice-then-regrain per role); the tracks
            # axis passes through unchanged.
            projected = {}
            p_start = None
            for role in self._value.roles:
                p = self.at(track=role).at(grain)
                projected[role] = p._value
                if p_start is None:
                    p_start = p._start
            out = Variable(display_name=self._display_name)
            out._value = TrackValues(projected)
            out.var_type = (
                'list' if out._value.time_length is not None else 'scalar'
            )
            out._unit = self._unit
            out._grain = grain
            out._start = p_start
            return out
        if self._indexed_by:
            # §17 rank-1 guard: without it, ``time`` resolving to None
            # would silently COPY the axised variable through the lens —
            # and downstream arithmetic with genuinely re-grained series
            # then misaligns coordinates against periods.
            raise ValueError(
                f"{label!r} is laid out along a finite axis — grain "
                f"projection of axised Variables lands with the track "
                f"layer. Project a slice (drop the axis first), or keep "
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
        spec = self._regrain
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
                "ratio and callable re-grain rules are evaluated by the model "
                "projection (a later slice); use a named recipe for now."
            )
        from .time import bucket_ranges
        n = len(self._value) if isinstance(self._value, list) else 1
        ranges = bucket_ranges(loc.grain, grain, loc.start, n)
        labels, values = regrain_series(self._value, loc.grain, grain, loc.start, rule)
        result = Variable(values, display_name=self._display_name,
                          start=labels[0], grain=grain, regrain=spec)
        # Live formula: each target cell renders as a reducer over the source's
        # native cells (=SUM(range) / period-end cell), falling back to the
        # inlined value when the source has no address in the workbook.
        result._set_expr(Regrain(source=self, buckets=[(lo, hi) for _, lo, hi in ranges],
                                 recipe=rule, fill_values=list(values)))
        return result

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

        Today this returns ``self._value`` directly. Exposed as a
        property (not a bare attribute) so future lazy-loading sources —
        DB-backed Variables, remote fetch, on-demand evaluation through
        an alternate compute engine — can hook here without touching
        callers. External code (user scripts, serializers, inspection
        tools) should use ``variable.value``. Internal engine code
        continues to use ``_value`` for directness.
        """
        return self._value

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
        intermediates leave it ``None``. Inspection tools — graph
        panels, formula-display UIs — read it via this property when
        they want the concrete call shape rather than the AST view.
        Returns ``None`` when no call-site snapshot was captured.
        """
        return self._source_code

    @source_code.setter
    def source_code(self, value: Optional[str]) -> None:
        """Override the call-site rendering — mostly used by extension
        authors building custom helpers. Setting to ``None`` clears it.
        """
        self._source_code = None if value is None else str(value)

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
        self._expr = value
        if value is None:
            self._dependency_refs = []
            return
        seen = set()
        refs: List['Variable'] = []
        for var in value.iter_refs():
            if id(var) not in seen:
                seen.add(id(var))
                refs.append(var)
        self._dependency_refs = refs

    def _set_expr(self, expr: Expr) -> None:
        """Thin delegate to the :attr:`expr` setter. Kept for the many
        existing call sites (``x._set_expr(node)``) that would otherwise
        need a mechanical rewrite. New code should use ``x.expr = node``.
        """
        self.expr = expr

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
        result._keys = list(self._keys) if self._keys is not None else None
        result._keys_source = self._keys_source
        result._excel_props = dict(self._excel_props)
        result._indexed_by = self._indexed_by
        # Time location + re-grain rule must survive a copy (adoption clones via
        # copy, so losing these would silently strip a line's grain / rule).
        result._start = self._start
        result._grain = self._grain
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
        return None
    
    @property
    def unit(self):
        """The measurement unit of this variable, or None."""
        return self._unit
    
    @unit.setter
    def unit(self, value: "str | None") -> None:
        self._unit = self._resolve_unit(value)
    
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
            fill_value = fill_value._value

        if periods > 0:
            shifted = [fill_value] * periods + self._value[:-periods]
        elif periods < 0:
            n = -periods
            shifted = self._value[n:] + [fill_value] * n
        else:
            shifted = list(self._value)

        # Build AST: shift(periods)  OR  shift(periods, fill_value=<...>)
        # Store the raw, typed fill value (already unwrapped from any
        # Variable above) so the renderer emits a correct Excel literal —
        # ``DATE(...)`` for dates, a bare number for numerics. Stringifying
        # here would lose the type and mis-render (quoted text / bare date).
        kwargs: Dict[str, Any] = {}
        if not (isinstance(original_fill_value, float) and original_fill_value == 0.0):
            kwargs["fill_value"] = fill_value

        result = Variable()
        result._set_expr(MethodCall(VarRef(self), "shift", [periods], kwargs))
        result._value = shifted
        result.var_type = 'list'
        result._keys = list(self._keys) if self._keys is not None else None

        return result

    # ═══════════════════════════════════════════════════════
    #  SPECIAL METHODS
    # ═══════════════════════════════════════════════════════
    
    def __index__(self) -> int:
        """Allow Variable to be used where Python expects an integer (range(), slicing, etc.).
        Only works for scalar Variables with integer-compatible values."""
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
        v = self._value
        if isinstance(v, list):
            v = v[0] if len(v) == 1 else v
        return int(v)

    def __float__(self) -> float:
        v = self._value
        if isinstance(v, list):
            v = v[0] if len(v) == 1 else v
        return float(v)

    def __bool__(self) -> bool:
        """Truthiness of the computed value — Python-mode branching works.

        ``if seats > 300:`` branches on the comparison's eager value
        (comparisons compute at operator time, per ADR-010), so plain
        Python control flow behaves exactly as Python users expect.
        Note the mode difference: a Python ``if`` decides NOW, in
        Python — the condition does not travel into the emitted Excel
        workbook. Use ``mo.IF(cond, then, otherwise)`` when the
        condition should live in the model (and in Excel) as a formula.

        Before this method existed, ``bool()`` fell back to ``__len__``
        (1 for scalars) — every ``if var:`` branch ran regardless of
        the value, silently. Value-truthiness fixes that.

        List-valued Variables refuse, pandas/numpy-style: the truth of
        many values at once is genuinely ambiguous.
        """
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
        indices on a scalar raise ``IndexError``.
        """
        if not isinstance(self._value, list):
            if key == 0:
                return Variable(value=self._value, var_type=self.var_type, value_type=self.value_type)
            raise IndexError(
                f"Index {key} on a scalar Variable — only ``[0]`` is valid "
                f"(as a convenience for list-shaped APIs). "
                f"Use ``var.value`` to read the scalar directly."
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

