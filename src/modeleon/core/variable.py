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

from typing import Optional, Set, Any, List, Dict, TYPE_CHECKING

from .humanize import humanize_identifier
from .component import Component
from .expr import Expr, MethodCall, Paren, Subscript, VarRef
from .qpath import QPath
from .variable_init import _VariableInit
from .variable_ops import _VariableArithmetic

if TYPE_CHECKING:
    from .multi_variable import MultiVariableBase


class Variable(_VariableInit, _VariableArithmetic, Component):
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
    })

    def __init__(self, value=None, value_type: str = "float", var_type: str = "scalar",
                 display_name: Optional[str] = None, formula=None, pyformula=None, unit=None, keys=None,
                 excel_props=None, **kwargs):
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

        # --- identity + default fields ---
        # ``_qualified_id``, ``_python_name``, ``_display_name``, and
        # ``_excel_props`` are initialized on the :class:`Component`
        # parent (with excel_props validated against ``_EXCEL_PROP_KEYS``).
        # The ``.path`` property below synthesizes a floating-namespace
        # path from ``id(self)`` until adoption crystallizes a rooted
        # path.
        super().__init__(display_name=display_name, excel_props=excel_props)
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

        self._process_kwargs(kwargs)

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
    
    # ``python_name`` property + setter inherited from :class:`Base`.
    # ``excel_props`` property, ``set_style``, ``set_display_name``,
    # ``set_python_name`` inherited from :class:`Component`.

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
        kwargs: Dict[str, Any] = {}
        if original_fill_value != 0.0:
            if isinstance(original_fill_value, Variable):
                # Preserve original behavior: use name-or-temp string (not VarRef)
                fill_value_repr = getattr(original_fill_value, 'python_name', None) or str(fill_value)
                kwargs["fill_value"] = fill_value_repr
            else:
                kwargs["fill_value"] = str(original_fill_value)

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

