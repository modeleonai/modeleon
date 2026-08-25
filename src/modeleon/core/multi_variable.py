# SPDX-License-Identifier: Apache-2.0
"""MultiVariable — the tree primitive.

A MultiVariable is a named container for Variables and nested MVs. The
tree of MVs rooted at a top-level ``Model`` is the source of truth for
model structure; the writer walks it to emit Excel tabs and cells.

Class cascade (in this file, top-to-bottom):

    MultiVariableBase       — internal base: state (components, _parent,
                              _python_name, _qualified_id) + identity
                              properties (python_name, path, id,
                              display_name) + attribute protocol
                              (__getattr__, __setattr__) + repr +
                              to_excel hook.
    MultiVariableClass      — reusable template; subclasses define
                              ``compute(**params)`` and Variables appear
                              as ``self.xxx`` assignments. Used for
                              cohort schedules, working-capital blocks,
                              and any other pattern you'd instantiate
                              multiple times.
    MultiVariable           — universal container for direct use.
                              ``excel_props={'tab': True}`` marks an Excel tab;
                              omit ``role`` for a plain grouping MV.

Three construction styles users can write::

    # Factory — keyword args become components
    pnl = MultiVariable("P&L", excel_props={'tab': True},
        revenue=Variable(1_000_000),
        cogs=Variable(600_000),
    )

    # Context manager — local names become component names
    with MultiVariable("P&L", excel_props={'tab': True}) as pnl:
        revenue = Variable(1_000_000)
        cogs = revenue * 0.6

    # Class-based — reusable template
    class CohortUnit(MultiVariableClass):
        def compute(self, start_users=100, churn=0.05):
            self.users = recurrence(start_users, "{prev} * (1 - {churn})",
                                     churn=churn)

Split across three mixin modules:

- :mod:`modeleon.core.mv_context` — ``_MVLifecycle``: start/end frame
  introspection, ``_register_component``, ``_clone``, the per-thread
  context stack.
- :mod:`modeleon.core.mv_inspect` — ``_MVInspect``: ``walk``,
  ``_collect_variables``, cycle detection, ``to_dict``, and other
  read-only tree navigation.
"""

from typing import Dict, List, Optional, Any, TYPE_CHECKING
import contextvars
import functools
import inspect

# Re-exported so code that imports these from modeleon.multi_variable keeps working.
from .component import Component
from .mv_context import _MVLifecycle
from .mv_inspect import _MVInspect
from .qpath import QPath

if TYPE_CHECKING:
    from pathlib import Path

# Scopes ``MultiVariableClass.__shell__`` construction: while set, a
# subclass ``compute()`` is not invoked by ``__init__``. A ContextVar
# (not a plain module flag) so suppression can't leak across threads
# or into constructors evaluated in argument position.
_SUPPRESS_COMPUTE: contextvars.ContextVar[bool] = contextvars.ContextVar(
    "mvclass_suppress_compute", default=False
)



@functools.lru_cache(maxsize=256)
def _signature_of(compute_func: Any) -> "inspect.Signature":
    """``inspect.signature`` keyed by the class's compute function."""
    return inspect.signature(compute_func)


class MultiVariableBase(_MVLifecycle, _MVInspect, Component):
    """
    Base class with internal machinery for grouped Variables and nested MultiVariables.
    
    This is the internal implementation - users should use:
    - MultiVariable: for quick one-off instances
    - MultiVariableClass: for reusable templates
    """
    
    # Default layout role. Users override per-instance via
    # ``MultiVariable(..., excel_props={'tab': True})`` to mark an Excel tab. The
    # writer dispatches on ``_role``, so new roles slot in additively
    # without touching existing code.
    _role: str = 'group'

    # Valid keys inside ``excel_props``. Passing any other key raises.
    _EXCEL_PROP_KEYS = frozenset({
        'hidden',   # служебное поддерево — скрыто во всех вьюхах
        # Structural
        'tab',
        # Section orientation: 'columns' spreads repeating children
        # ACROSS (instances as columns, scalar fields as rows).
        'orient',
        # The section's article number («1.2») — see Variable.
        'article', 'row', 'col',
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

    def __init__(
        self,
        display_name: Optional[str] = None,
        excel_props: Optional[Dict[str, Any]] = None,
        excel_layout: Optional[Any] = None,
        description: Optional[str] = None,
        **kwargs: Any,
    ):
        """
        Initialize MultiVariableBase.

        Args:
            display_name: User-friendly label (e.g. "Income Statement").
                   If not set, auto-derived from python_name.
            excel_props: Dict of Excel-renderer hints for this MV.
                   Recognized keys: ``'tab'`` (mark as an Excel tab),
                   ``'row'`` / ``'col'`` (explicit placement), plus
                   styling (``'bold'``, ``'bg'``, ``'number_format'``, …
                   see :attr:`_EXCEL_PROP_KEYS`). Unknown keys raise.
            excel_layout: Optional opaque layout-config object. The
                   engine stores it as ``self._excel_layout`` and does
                   nothing else with it; the writer / pro layer reads
                   it back. Designed for pro's ``ExcelLayout`` dataclass
                   (template name + sheet/address/format overrides) but
                   engine stays duck-typed — anything that round-trips
                   through ``repr()`` is fine. ``None`` means inherit
                   from parent at render time. See
                   ``modeleon_pro.excel_layout`` for the dataclass shape.

        ``MultiVariableBase`` itself accepts only ``display_name``,
        ``excel_props``, and ``excel_layout``. The :class:`MultiVariable`
        subclass is what users instantiate; it absorbs Variable / MV-
        typed kwargs as components and forwards the rest here. Anything
        that lands in ``kwargs`` at this layer is a typo and raises
        ``TypeError`` — mirrors :class:`Variable`'s strict-kwarg
        discipline.
        """
        if kwargs:
            raise TypeError(
                f"{type(self).__name__}() got unexpected keyword argument(s): "
                f"{', '.join(sorted(kwargs))}. "
                f"Known kwargs: display_name, description, excel_props, excel_layout. "
                f"For component children, assign attributes "
                f"(``mv.revenue = mo.Variable(...)``) or use the factory form "
                f"(``MultiVariable(revenue=mo.Variable(...))``). "
                f"For styling and structural placement, pass "
                f"``excel_props={{'row': N, 'col': N, 'bold': True, ...}}``. "
                f"For top-level naming, use ``mo.Model('name')``."
            )

        # ``_qualified_id``, ``_python_name``, ``_display_name``, and
        # ``_excel_props`` are initialized on :class:`Component`; the
        # excel_props dict is validated against ``_EXCEL_PROP_KEYS`` there.
        # Adoption (``parent.child = mv`` → ``_register_component``) is
        # what later sets ``_python_name``.
        super().__init__(
            display_name=display_name,
            excel_props=excel_props,
            description=description,
        )

        # Opaque layout-config attribute. Pro reads this back via the
        # writer; engine never inspects its shape. ``None`` is the
        # inherit-from-parent signal in pro's resolution walker.
        self._excel_layout = excel_layout

        # Component name registry. Components themselves (Variables and
        # nested MultiVariables) live in ``self.__dict__`` directly so
        # they're accessible via normal Python attribute lookup AND
        # via ``self.__dict__`` (which a future synthetic-module
        # rendering path can exec into). ``_component_names`` is an
        # explicit list of which ``__dict__`` keys are components —
        # disambiguating from private state attrs (``_parent``,
        # ``_qualified_id``, etc.) and from Variable IDs whose binding
        # names start with underscore (floating qpaths).
        self._component_names: List[str] = []

        # Layout role is a derived internal marker. Users set
        # ``excel_props={'tab': True}`` to mark an Excel tab; the writer
        # and layout engine read ``_role`` / ``_is_sheet`` from there.
        if self._excel_props.get('tab'):
            self._role = 'sheet'
        else:
            self._role = type(self)._role

        # Parent-child tracking
        self._parent: Optional['MultiVariableBase'] = None
        self._name_in_parent: Optional[str] = None

    def _set_default_window(
        self,
        default_grain: Optional[str] = None,
        default_start: Optional[str] = None,
        default_periods: Optional[Any] = None,
    ) -> None:
        """Validate and set the ambient time window (``default_grain`` /
        ``default_start`` / ``default_periods``) declared in a constructor.

        The window is where a model's time lives: list-valued lines inherit
        their native ``(start, grain)`` from the nearest ancestor window
        (``Variable.time``), and the Excel/spreadsheet timeline header is
        derived from it (``resolve_default_window``). Plain attribute
        assignment (``m.default_grain = 'month'``) remains valid and
        unvalidated; this constructor path fails loudly on a bad window.
        """
        from .time import GRAINS, _parse
        if default_grain is None:
            given = 'default_start' if default_start is not None else 'default_periods'
            raise ValueError(
                f"{given}= needs default_grain= beside it — the window's grain "
                f"is what anchors the start label and the period count "
                f"(e.g. default_grain='month', default_start='2026-01', "
                f"default_periods=24)."
            )
        if default_grain not in GRAINS:
            raise ValueError(
                f"Unknown default_grain {default_grain!r}; supported: "
                f"{', '.join(GRAINS)}."
            )
        if default_start is not None:
            _parse(default_start, default_grain)     # raises with the format hint
        if default_periods is not None:
            if isinstance(default_periods, bool) or not isinstance(default_periods, int):
                raise TypeError(
                    f"default_periods must be an int; got "
                    f"{type(default_periods).__name__}."
                )
            if default_periods <= 0:
                raise ValueError(
                    f"default_periods must be positive; got {default_periods}."
                )
        self.default_grain = default_grain
        if default_start is not None:
            self.default_start = default_start
        if default_periods is not None:
            self.default_periods = default_periods

    @property
    def _components(self) -> Dict[str, Any]:
        """Dict view of components — derived from ``__dict__`` indexed
        by ``_component_names``. Read returns a fresh dict snapshot."""
        return {k: self.__dict__[k] for k in self._component_names if k in self.__dict__}

    @_components.setter
    def _components(self, mapping: Dict[str, Any]) -> None:
        """Bulk-replace components. Used by virtual-MV rendering paths.
        Wipes the old set (removes from ``__dict__`` and clears
        ``_component_names``), installs the new ones."""
        current = list(self.__dict__.get('_component_names', []))
        for name in current:
            self.__dict__.pop(name, None)
        self.__dict__['_component_names'] = list(mapping.keys())
        for name, value in mapping.items():
            self.__dict__[name] = value

    @property
    def _component_order(self) -> List[str]:
        """Insertion-ordered component names."""
        return list(self._component_names)

    @_component_order.setter
    def _component_order(self, order: List[str]) -> None:
        """Replace the component name order. Components themselves stay
        in ``__dict__``; only the order list is updated."""
        self.__dict__['_component_names'] = list(order)

    @property
    def _is_sheet(self) -> bool:
        """Whether this MultiVariable is an Excel tab.

        Derived from ``_role``: pass ``excel_props={'tab': True}`` to
        :class:`MultiVariable` to create a tab.
        """
        return self._role == 'sheet'

    # Context-manager lifecycle (start, end, __enter__, __exit__) and
    # component adoption (_register_component, _clone) live on
    # :class:`_MVLifecycle` in ``mv_context.py``. MultiVariableBase
    # inherits them via the class declaration above.

    def to_excel(self, path: "str | Path") -> None:
        """Emit this MV (and everything inside it) to an ``.xlsx`` file.

        Every MultiVariable can emit — no wrapper required.
        ``pnl.to_excel("pnl_only.xlsx")`` exports just the P&L subtree;
        ``model.to_excel("full.xlsx")`` exports the whole thing.
        """
        from ..compile.excel.writer import to_excel as _to_excel
        _to_excel(path, root=self)

    def at(self, grain: str) -> 'MultiVariableBase':
        """The model re-grained (projected) onto a coarser ``grain``. Every line
        re-grains by its own rule (apply-own-rule), and **formulas without a rule
        recompute over their re-grained inputs** (correct for additive / ratio
        formulas — only multiplicative flows need an explicit ``sum``). Returns a
        new ``MultiVariable`` of re-grained values, ready to inspect or
        ``.to_excel``.
        """
        from .projection import project_model
        return project_model(self, grain)

    def __getattr__(self, name: str):
        """Component access fallback.

        Components live in ``self.__dict__``, so normal Python attribute
        lookup finds them. ``__getattr__`` only runs for missing names.
        """
        if name.startswith('_'):
            raise AttributeError(f"'{type(self).__name__}' object has no attribute '{name}'")
        raise AttributeError(f"'{type(self).__name__}' has no component '{name}'")

    def __setattr__(self, name: str, value: Any) -> None:
        """Auto-register Variables or nested MVs assigned as attributes.

        ``model.pnl = mo.MultiVariable(...)`` and
        ``sheet.revenue = mo.Variable(...)`` flow through here and
        become ``self._register_component(name, value)``. This is the
        only path by which a Variable / MV joins the model.

        Private (underscore-prefixed) attrs bypass registration and go
        through normal attribute assignment.
        """
        from .variable import Variable
        if (
            not name.startswith('_')
            # ``default_excel_view`` is a presentation POINTER slot, not content.
            # Its value is an ``ExcelView`` (itself a MultiVariable), but it must
            # never join ``_components`` or it would render as a spurious tab.
            and name != 'default_excel_view'
            and '_component_names' in self.__dict__
            and (isinstance(value, Variable) or isinstance(value, MultiVariableBase))
        ):
            self._register_component(name, value)
            # ``_register_component`` may have substituted ``value`` with
            # a clone (cross-parent adoption). Use the registered version
            # so the attribute slot holds the same instance as
            # ``self.__dict__[name]``.
            value = self.__dict__.get(name, value)
        super().__setattr__(name, value)
    
    # ``python_name`` property + setter are inherited from :class:`Base`.

    @property
    def path(self) -> QPath:
        """Qualified path identity — the single source of identity.

        Resolution order on an un-crystallized MV:

        1. If ``_parent`` is set and its path is rooted, derive from
           parent + ``_name_in_parent`` and cache.
        2. Else if ``python_name`` is set AND there's no parent
           (top-level MV with an explicit binding), crystallize to
           ``ROOT.child(name)`` and cascade into any components
           adopted before now.
        3. Else return a floating path derived from ``id(self)``.
        """
        if self._qualified_id is not None:
            return self._qualified_id
        if self._parent is not None:
            parent_path = self._parent.path
            if not parent_path.is_floating and self._name_in_parent:
                self._qualified_id = parent_path.child(self._name_in_parent)
                return self._qualified_id
        elif self._python_name:
            self._qualified_id = QPath.ROOT.child(self._python_name)
            from .mv_context import _crystallize_subtree
            for comp_name in self._component_order:
                child = self._components[comp_name]
                _crystallize_subtree(child, self._qualified_id.child(comp_name))
            return self._qualified_id
        return QPath.floating(id(self), kind="m")

    # ``id`` property is inherited from :class:`Base`.

    @property
    def display_name(self) -> str:
        """User-friendly display name.

        Resolution order:
        1. ``_display_name`` — explicit ``display_name=`` kwarg.
        2. Humanized Python identifier — ``income_statement`` →
           ``"Income Statement"``, ``XIRR`` → ``"XIRR"`` (all-caps
           preserved).

        Settable: ``mv.display_name = 'Custom Label'`` overrides the
        fallback. Set to ``None`` to revert.
        """
        if self._display_name:
            return self._display_name
        from .humanize import humanize_identifier
        base = self._name_in_parent or self.python_name or self.path.leaf
        return humanize_identifier(base)

    @display_name.setter
    def display_name(self, value: Optional[str]) -> None:
        # Setter mirrors Component.set_display_name; redeclared here so
        # the property + setter pair sits on the concrete class with the
        # custom getter above.
        self._display_name = value

    @classmethod
    def to_new_source(cls, name: str) -> str:
        """Source form of a NEW, empty instance bound to ``name`` —
        the class describes its own constructor spelling (the source
        sibling of ``__repr__``). Subclasses whose constructor takes
        the name positionally override (e.g. :class:`Model`).
        """
        return "mo.MultiVariable()"

    def add(self, name: str, component: "Component") -> "Component":
        """Adopt ``component`` as a named child under a runtime string
        ``name`` — the dynamic-name form of ``mv.<name> = component``.

        Equivalent to attribute assignment (routes through the same
        ``__setattr__`` → ``_register_component`` adoption), but lets the
        child name be computed at runtime — the one thing ``mv.attr =``
        can't express. Returns the adopted instance, which may be a
        clone when ``component`` was already parented elsewhere
        (clone-on-ownership-change). Re-using an existing ``name``
        orphans the previous child (with a ``ModelStructureWarning``),
        so callers adding in a loop must compute unique names.
        """
        if not isinstance(name, str) or not name or name.startswith("_"):
            raise ValueError(
                "mv.add(name, ...) requires a non-empty public identifier "
                f"(no leading underscore). Got {name!r}."
            )
        from .variable import Variable
        if not isinstance(component, (Variable, MultiVariableBase)):
            raise TypeError(
                "mv.add adopts Variables / MultiVariables only. Got "
                f"{type(component).__name__}."
            )
        setattr(self, name, component)
        return self.__dict__.get(name, component)

    def remove(self, name: str) -> "Component":
        """Detach and unregister the child component named ``name`` — the
        inverse of :meth:`add` / attribute assignment.

        Clears the child's parent/owner links and drops it from this
        MV's components, returning the detached instance. Raises
        :class:`KeyError` if there is no such component.
        """
        if not isinstance(name, str) or name not in self._components:
            raise KeyError(
                f"{name!r} is not a component of this "
                f"{type(self).__name__}."
            )
        return self._detach_component(name)

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}({self.id}, components={len(self._components)})"

    def _repr_html_(self) -> str:
        """Jupyter / IPython rich-display: render every MV as a tabbed view.

        Each sub-MultiVariable becomes a tab; direct Variables on
        ``self`` become a synthetic overview tab labeled with the MV's
        own name. A bare MV (only direct Variables, no sub-MVs) renders
        as a single tab — matching how the same MV would appear as a
        sheet in Excel.
        """
        from ..display.html import model_html
        return model_html(self)


class MultiVariableClass(MultiVariableBase):
    """
    Base class for reusable MultiVariable templates — cohort
    schedules, working-capital blocks, any pattern you'd instantiate
    multiple times with different parameters.

    Features:
    - Automatic value extraction from Variable parameters
    - Two coding styles: ``compute()`` method or ``__init__``

    Variable Parameters:
        Parameters can be raw values OR Variable objects. Values are extracted automatically:

        # Both work identically:
        unit = CohortUnit(start_users=100, churn_rate=0.05)

        start = Variable(100)
        rate  = Variable(0.05)
        unit = CohortUnit(start_users=start, churn_rate=rate)

    Style A — using ``compute()`` (zero boilerplate)::

        class CohortUnit(MultiVariableClass):
            def compute(self, start_users=100, churn_rate=0.05):
                self.users = mo.recurrence(start_users, "{prev} * (1 - {c})", c=churn_rate)

        unit = CohortUnit(start_users=500)   # just works

    Style B — using ``__init__`` (explicit, familiar)::

        class CohortUnit(MultiVariableClass):
            def __init__(self, start_users=100, churn_rate=0.05):
                super().__init__()
                self.users = mo.recurrence(start_users, "{prev} * (1 - {c})", c=churn_rate)

        unit = CohortUnit(start_users=500)   # also works
    """
    
    def __init__(self, **kwargs):
        """
        Initialize MultiVariableClass.

        Args:
            **kwargs: Parameters passed to compute() or stored for later use.
                      Can be raw values or Variable objects (values will be extracted).
                      Special key: display_name (forwarded to MultiVariableBase).
        """
        from .variable import Variable
        self._Variable_class = Variable

        # Pop the kwargs the parent ``MultiVariableBase`` recognizes;
        # everything else is a template parameter (``start_users=100``,
        # ``churn_rate=0.05``) for ``compute()`` or attribute storage.
        display_name = kwargs.pop('display_name', None)
        excel_props = kwargs.pop('excel_props', None)
        excel_layout = kwargs.pop('excel_layout', None)

        super().__init__(
            display_name=display_name,
            excel_props=excel_props,
            excel_layout=excel_layout,
        )

        self._input_variables: Dict[str, 'Variable'] = {}

        for key, value in kwargs.items():
            extracted = self._extract_value(value)
            setattr(self, key, extracted)
            if isinstance(value, Variable):
                self._input_variables[key] = value

        # If compute() is overridden by subclass, call it with resolved
        # params — unless a ``__shell__`` construction asked for the
        # container empty (children to be populated explicitly).
        self._compute_suppressed = _SUPPRESS_COMPUTE.get()
        if (
            type(self).compute is not MultiVariableClass.compute
            and not self._compute_suppressed
        ):
            self._call_compute()

    @classmethod
    def __shell__(cls, **kwargs):
        """Construct an instance WITHOUT invoking ``compute()``.

        Returns a real instance of ``cls`` — ``isinstance`` checks,
        methods, properties, and the MRO all behave normally — but the
        container starts empty: constructor kwargs are stored exactly
        as in normal construction (attributes + ``_input_variables``),
        and children are expected to be assigned explicitly afterwards.

        For tooling that re-creates a previously computed container
        from its recorded statements (deserialization, code
        generation) without running the template body a second time::

            u = CohortUnit.__shell__(start_users=500)
            u.users = ...      # children populated by the caller

        Suppression is scoped by a :class:`~contextvars.ContextVar`,
        so classes constructed in argument position (``__shell__(sub=
        Other(...))``) still compute normally.
        """
        token = _SUPPRESS_COMPUTE.set(True)
        try:
            return cls(**kwargs)
        finally:
            _SUPPRESS_COMPUTE.reset(token)
    
    def _call_compute(self) -> None:
        """Introspect compute() signature and call with resolved params.
        For parameters that were originally Variable objects, pass the Variable
        (not the extracted scalar) so that formula tracking is preserved.

        Required parameters (no default) resolve from the constructor
        kwargs stored by ``__init__`` — ``Unit(seats=100)`` must reach
        ``compute(self, seats)``. A required parameter with no stored
        value is left out so ``compute()`` raises its natural TypeError
        naming the missing argument."""
        # Signature by UNDERBOUND function, cached: a register holds
        # hundreds of instances of the same four classes, and
        # ``inspect.signature`` costs ~7 мкс a call — thousands of
        # identical introspections per run for four distinct answers.
        sig = _signature_of(type(self).compute)
        params = {}
        for k, v in sig.parameters.items():
            if v.kind in (
                inspect.Parameter.VAR_POSITIONAL,
                inspect.Parameter.VAR_KEYWORD,
            ):
                continue
            if k in self._input_variables:
                params[k] = self._input_variables[k]
            elif hasattr(self, k):
                params[k] = getattr(self, k)
            elif v.default is not inspect.Parameter.empty:
                params[k] = v.default
        self.compute(**params)
    
    def _extract_value(self, value: Any) -> Any:
        """
        Extract raw value from a Variable, or return as-is.

        This enables passing Variables as parameters to compute()::

            churn = Variable(0.05)
            unit = CohortUnit(churn_rate=churn)   # Works!
        """
        if isinstance(value, self._Variable_class):
            return value._value
        return value
    
    def compute(self, **kwargs):
        """
        Override to create components. Assign Variables directly:
        ``self.users = mo.Variable(...)``.

        Parameters are auto-resolved from:
        1. Instance attributes (set via __init__ kwargs)
        2. Default values in method signature
        """
        pass

    # __setattr__ for imperative-style component registration is inherited
    # from MultiVariableBase — no class-specific override needed.


class MultiVariable(MultiVariableBase):
    """
    Universal container for Variables and nested MultiVariables.

    Supports three construction styles:

    Style 1: Imperative (attributes assigned after creation)
        income = MultiVariable("Income Statement", excel_props={'tab': True})
        income.revenue = Variable(1_000_000)
        income.cogs = income.revenue * Variable(0.6)

    Style 2: Factory (keyword arguments)
        assumptions = MultiVariable(
            "Assumptions", excel_props={'tab': True},
            tax_rate=Variable(0.25),
            cogs_pct=Variable(0.60),
        )

    Style 3: Context manager — local names become component names
        with MultiVariable("P&L", excel_props={'tab': True}) as pnl:
            revenue = Variable(1_000_000)
            cogs = revenue * 0.6

    Pass ``excel_props={'tab': True}`` to mark an Excel tab; omit ``role`` for a
    plain grouping MV.
    """

    def __init__(self, display_name=None, **components):
        """
        Create a MultiVariable.

        Args:
            display_name: User-friendly label (e.g. "Income Statement").
            **components: Named Variables or MultiVariables to include,
                plus optional ``excel_props={'tab': True}`` to mark an
                Excel tab.
        """
        from .variable import Variable

        # Carve out the kwargs ``MultiVariableBase`` recognizes by name
        # so they're forwarded explicitly rather than treated as
        # candidate components.
        base_kwargs: Dict[str, Any] = {}
        if "excel_props" in components:
            base_kwargs["excel_props"] = components.pop("excel_props")
        if "excel_layout" in components:
            base_kwargs["excel_layout"] = components.pop("excel_layout")
        if "description" in components:
            base_kwargs["description"] = components.pop("description")

        # The ambient time window — declared once on a container, inherited
        # by every list-valued line under it (``Variable.time``) and by the
        # Excel timeline header (``resolve_default_window``).
        window_kwargs: Dict[str, Any] = {}
        for window_key in ("default_grain", "default_start", "default_periods"):
            if window_key in components:
                window_kwargs[window_key] = components.pop(window_key)

        # The tracks axis (§16.2) — declared once, inherited ambiently
        # exactly like the window (``resolve_tracks_decl``). Role-kwarg
        # Variables (``actual=``/``plan=``) materialize against it at
        # adoption.
        # The VIEW slot, ctor spelling. ``__setattr__`` exempts
        # ``default_excel_view`` from component registration — but the
        # ctor's registerable loop calls ``_register_component``
        # DIRECTLY, so a view passed as a kwarg was adopted as CONTENT:
        # its config leaves (``totals=['quarter','year']``) hit the
        # extent law under a windowed model («'totals' has 2 value(s)
        # in a 36-period window», live on a focused Бизнес-план) and
        # the view itself would render as a spurious tab.
        view_slot = components.pop("default_excel_view", None)

        tracks_decl = components.pop("tracks", None)
        if tracks_decl is not None:
            from .tracks_decl import Tracks
            if not isinstance(tracks_decl, Tracks):
                raise TypeError(
                    f"tracks= takes a mo.Tracks declaration, e.g. "
                    f"tracks=mo.Tracks('факт', 'бюджет'); got "
                    f"{type(tracks_decl).__name__}."
                )

        registerable = {}
        other_kwargs = {}
        for comp_name, value in components.items():
            if isinstance(value, (Variable, MultiVariableBase)):
                registerable[comp_name] = value
            else:
                other_kwargs[comp_name] = value

        super().__init__(display_name=display_name, **base_kwargs, **other_kwargs)

        if window_kwargs:
            self._set_default_window(**window_kwargs)

        if tracks_decl is not None:
            # Plain attribute (Tracks is neither Variable nor MV, so no
            # component registration) — ``resolve_tracks_decl`` reads it
            # via getattr on the ancestor walk.
            self.tracks = tracks_decl
            if tracks_decl.blend is not None:
                from .time import resolve_default_window
                w = resolve_default_window(self)
                if w is None or w.grain is None or w.start is None \
                        or w.periods is None:
                    raise ValueError(
                        "mo.blend needs the fully declared window beside "
                        "it — the boundary is a date, so the declaration "
                        "needs default_grain, default_start and "
                        "default_periods (otherwise the synthesized "
                        "track would silently never appear)."
                    )

        if view_slot is not None:
            # Through the exempted slot — a pointer, never content.
            self.default_excel_view = view_slot

        for name, comp in registerable.items():
            self._register_component(name, comp)


