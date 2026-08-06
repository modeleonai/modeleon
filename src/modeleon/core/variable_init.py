# SPDX-License-Identifier: Apache-2.0
"""Construction dispatch for :class:`Variable`.

Pure "turn a kwargs bundle into Variable state" logic. Every method
here runs once during ``Variable.__init__`` and is never consulted
afterward — so it doesn't belong in the main ``variable.py`` alongside
the Variable's runtime behavior.

The mixin :class:`_VariableInit` is inherited by Variable. Methods
mutate ``self`` directly (setting ``_value``, ``_expr``,
``_dependency_refs``, ``pyformula``, ``_raw_formula_str``, ``_styles``,
``_source_code``) — they're not pure functions, just isolated ones.
"""

from __future__ import annotations

from typing import Any, Dict, FrozenSet, List, Optional, TYPE_CHECKING

from .expr import Expr, ListExpr, Literal, VarRef

if TYPE_CHECKING:
    from .variable import Variable


def _plugin_variable_kwargs() -> set:
    """Return the set of extra Variable kwargs registered by plugins.

    Called per-Variable construction so plugins registered after this
    module was first imported still take effect. Empty set when no
    plugins are loaded.
    """
    from ..plugins import get_registry
    return set(get_registry().get("variable_kwargs", {}).keys())


class _VariableInit:
    """Construction mixin: dispatchers for value / formula / pyformula,
    type coercion, auto-detection, and kwarg sorting.

    Attributes/methods listed below are provided by the concrete
    :class:`Variable` subclass — declared here for type-checkers.
    """

    # Attributes provided by Variable. Listed here so writes inside this
    # mixin (e.g. ``self._source_code = ...``) line up with their final
    # types on the concrete class.
    _EXCEL_PROP_KEYS: FrozenSet[str]
    pyformula: Optional[str]
    _source_code: Optional[str]
    _value: Any
    _expr: Optional[Expr]
    _raw_formula_str: Optional[str]
    _dependency_refs: List["Variable"]

    if TYPE_CHECKING:
        def _set_expr(self, expr: Expr) -> None: ...

    @staticmethod
    def _resolve_keys(keys) -> tuple[Optional[list], Optional["Variable"]]:
        """Normalize the ``keys`` kwarg into (keys_list, keys_source)."""
        from .variable import Variable
        if keys is None:
            return None, None
        if isinstance(keys, Variable):
            keys_list = list(keys._value) if isinstance(keys._value, list) else [keys._value]
            return keys_list, keys
        if isinstance(keys, list):
            return list(keys), None
        raise TypeError(
            f"keys= must be a list of labels or a Variable whose values "
            f"become the labels, got {type(keys).__name__}. "
            f"Examples:\n"
            f"  Variable([800, 1000, 1200], keys=['Bear', 'Base', 'Bull'])\n"
            f"  Variable([10, 20, 30], keys=years_variable)\n"
            f"Or use the dict-sugar form: ``Variable({{'Bear': 800, ...}})``."
        )

    def _init_from_formula(self, formula, var_type: str) -> str:
        """Handle the ``formula=`` argument. Returns the updated var_type."""
        from .variable import Variable
        if isinstance(formula, Variable):
            # Copy another Variable's expression into this one (by reference)
            self._value = formula._value
            var_type = formula.var_type
            if formula._expr is not None:
                self._set_expr(formula._expr)
            elif formula.is_formula:
                # Source has a raw-string formula with no AST — preserve as string
                self._raw_formula_str = formula.formula
                self._dependency_refs = list(formula._dependency_refs)
            else:
                self._set_expr(VarRef(formula))
            if formula._source_code is not None:
                self._source_code = formula._source_code
            return var_type

        if isinstance(formula, list) and len(formula) > 0:
            # Mixed Variable/scalar list → ListExpr compound formula
            self._build_compound_formula(formula)
            return 'list'

        if isinstance(formula, Expr):
            # Direct AST construction — preferred path from operators and helpers
            self._set_expr(formula)
            self._value = None
            return var_type

        # Plain string formula (rare — user-provided raw strings)
        self._raw_formula_str = str(formula)
        self._value = None
        return var_type

    def _init_from_pyformula(self, pyformula, var_type: str) -> str:
        """Handle the ``pyformula=`` argument. Returns the updated var_type."""
        from .variable import Variable
        if isinstance(pyformula, Variable):
            self._value = pyformula._value
            var_type = pyformula.var_type
            self._dependency_refs = [pyformula]
            self.pyformula = pyformula.python_name or pyformula.path.leaf
        else:
            # Any Python value — captured as string for code generation
            self._value = pyformula
            self.pyformula = str(pyformula)
        return var_type

    def _init_from_value(self, value, var_type: str, value_type: str) -> tuple[str, str]:
        """Handle the ``value=`` argument. Returns (var_type, value_type)."""
        from .variable import Variable
        from .tracks import TrackValues
        if isinstance(value, TrackValues):
            # A tracked value (role → series) stores as-is; var_type
            # follows the tracks' shape (list tracks → a series row).
            self._value = value
            return ('list' if value.time_length is not None else 'scalar',
                    value_type)
        if isinstance(value, Variable):
            # A passed-in Variable is a reference, not a copy of the inputs
            self._value = value._value
            var_type = value.var_type
            if value._expr is not None:
                self._set_expr(value._expr)
            elif value.is_formula:
                self._raw_formula_str = value.formula
                self._dependency_refs = list(value._dependency_refs)
            else:
                self._set_expr(VarRef(value))
            if value._source_code is not None:
                self._source_code = value._source_code
            return var_type, value_type

        if (isinstance(value, list) and len(value) > 0
                and any(isinstance(v, Variable) for v in value)):
            # List with Variables → ListExpr compound formula
            self._build_compound_formula(value)
            return 'list', value_type

        # Plain literal: scalar, list of scalars, date, string, etc.
        self._value = self._coerce_literal(value, value_type)
        var_type, value_type = self._auto_detect_types(self._value, var_type, value_type)
        return var_type, value_type

    def _build_compound_formula(self, items: list) -> None:
        """Build ``[id1, 0, id2, 5]``-style compound formula from a list that
        contains Variables and/or scalars. Used by both formula-mode
        (``formula=`` is a list) and value-mode (``value=`` is a list
        containing Variables)."""
        from .variable import Variable
        self._value = []
        expr_items: List[Expr] = []
        for item in items:
            if isinstance(item, Variable):
                self._value.append(item._value if item._value is not None else None)
                expr_items.append(VarRef(item))
            else:
                self._value.append(item)
                expr_items.append(Literal(item))
        self._set_expr(ListExpr(expr_items))

        # Copy _source_code if every Variable item shares the same one
        var_items = [v for v in items if isinstance(v, Variable)]
        if var_items and all(v._source_code for v in var_items):
            # ``all(...)`` above guarantees every entry is a non-empty str.
            source_codes: List[str] = [v._source_code for v in var_items]  # type: ignore[misc]
            if len(set(source_codes)) == 1:
                self._source_code = source_codes[0]

    @staticmethod
    def _coerce_literal(value, value_type: str):
        """Parse ISO date strings when the user explicitly asks for datetime;
        otherwise pass the value through unchanged."""
        from datetime import datetime as dt
        if value_type not in ('datetime', 'date'):
            return value
        if isinstance(value, str):
            return dt.strptime(value, '%Y-%m-%d').date()
        if isinstance(value, list):
            return [dt.strptime(v, '%Y-%m-%d').date() if isinstance(v, str) else v for v in value]
        return value

    @staticmethod
    def _auto_detect_types(value, var_type: str, value_type: str) -> tuple[str, str]:
        """Infer (var_type, value_type) from the concrete value when the
        user left them at their defaults. Explicit kwargs are respected."""
        from datetime import date, datetime as dt

        if isinstance(value, list):
            if var_type == 'scalar':
                var_type = 'list'
            if value_type == 'float' and len(value) > 0:
                first = value[0]
                if isinstance(first, (date, dt)):
                    value_type = 'datetime'
                elif isinstance(first, int):
                    value_type = 'int'
                elif isinstance(first, str):
                    value_type = 'string'
        elif isinstance(value, (date, dt)):
            value_type = 'datetime'
        elif isinstance(value, int):
            value_type = 'int'
        elif isinstance(value, str):
            value_type = 'string'
        return var_type, value_type

    def _process_kwargs(self, kwargs: Dict[str, Any]) -> Dict[str, Any]:
        """Sort extra kwargs into plugin attributes; return the rest.

        Known kwargs that ``Variable.__init__`` accepts directly (value,
        formula, display_name, unit, keys, excel_props, …) are already
        consumed by ``__init__``'s signature. What lands in ``kwargs``
        is whatever else the user passed — plugin-registered extensions
        (``control=``, ``access=``), TRACK kwargs (§16.2: track names
        are user content, so they cannot be literal parameters), or
        typos. The remainder is returned for ``__init__`` to stash as
        track candidates; validation against the ambient ``mo.Tracks``
        declaration happens at adoption, where a name that is neither
        a declared track nor a known kwarg dies with a teaching error.
        """
        remainder: Dict[str, Any] = {}
        plugin_kwargs = _plugin_variable_kwargs()
        for k, v in kwargs.items():
            if k in plugin_kwargs:
                setattr(self, f"_{k}", v)
            else:
                remainder[k] = v
        return remainder

    # ``_process_excel_props`` consolidated into
    # ``Component._validate_excel_props`` and called via
    # ``super().__init__(..., excel_props=...)`` from Variable.__init__.
