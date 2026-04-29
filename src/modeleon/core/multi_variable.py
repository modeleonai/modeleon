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
import inspect

# Re-exported so code that imports these from modeleon.multi_variable keeps working.
from .component import Component
from .mv_context import _MVLifecycle
from .mv_inspect import _MVInspect
from .qpath import QPath

if TYPE_CHECKING:
    from pathlib import Path



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
        # Structural
        'tab', 'row', 'col',
        # Typography
        'bold', 'italic', 'font_color', 'font_size', 'font_family',
        'text_align', 'indent',
        # Fill / borders
        'bg', 'border_top', 'border_bottom', 'border_left', 'border_right',
        # Data formatting
        'number_format',
    })

    def __init__(
        self,
        display_name: Optional[str] = None,
        excel_props: Optional[Dict[str, Any]] = None,
        **kwargs
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
            **kwargs: Captured on ``self._creation_params`` for
                  inspection by downstream tooling.
        """
        # ``_qualified_id``, ``_python_name``, ``_display_name``, and
        # ``_excel_props`` are initialized on :class:`Component`; the
        # excel_props dict is validated against ``_EXCEL_PROP_KEYS`` there.
        # Adoption (``parent.child = mv`` → ``_register_component``) is
        # what later sets ``_python_name``.
        super().__init__(display_name=display_name, excel_props=excel_props)

        self._creation_params: Dict[str, Any] = kwargs.copy()

        # Component storage (Variables and nested MultiVariables)
        self._components: Dict[str, Any] = {}
        self._component_order: List[str] = []

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

    def __getattr__(self, name: str):
        """Access component (Variable or nested MultiVariable) by attribute name."""
        if name.startswith('_'):
            raise AttributeError(f"'{type(self).__name__}' object has no attribute '{name}'")

        if '_components' in self.__dict__ and name in self._components:
            return self._components[name]

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
            and '_components' in self.__dict__
            and (isinstance(value, Variable) or isinstance(value, MultiVariableBase))
        ):
            self._register_component(name, value)
            # ``_register_component`` may have substituted ``value`` with
            # a clone (cross-parent adoption). Use the registered version
            # so the attribute slot holds the same instance as
            # ``self._components[name]``.
            value = self._components.get(name, value)
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

        display_name = kwargs.pop('display_name', None)

        self._input_variables: Dict[str, 'Variable'] = {}

        for key, value in kwargs.items():
            extracted = self._extract_value(value)
            setattr(self, key, extracted)
            if isinstance(value, Variable):
                self._input_variables[key] = value

        base_kwargs = {k: getattr(self, k) for k in kwargs if hasattr(self, k)}
        super().__init__(display_name=display_name, **base_kwargs)

        # If compute() is overridden by subclass, call it with resolved params
        if type(self).compute is not MultiVariableClass.compute:
            self._call_compute()
    
    def _call_compute(self) -> None:
        """Introspect compute() signature and call with resolved params.
        For parameters that were originally Variable objects, pass the Variable
        (not the extracted scalar) so that formula tracking is preserved."""
        sig = inspect.signature(self.compute)
        params = {}
        for k, v in sig.parameters.items():
            if v.default is not inspect.Parameter.empty:
                if k in self._input_variables:
                    params[k] = self._input_variables[k]
                else:
                    params[k] = getattr(self, k, v.default)
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
                plus optional ``excel_props={'tab': True}`` to mark an Excel tab.
        """
        from .variable import Variable

        registerable = {}
        other_kwargs = {}
        for comp_name, value in components.items():
            if isinstance(value, (Variable, MultiVariableBase)):
                registerable[comp_name] = value
            else:
                other_kwargs[comp_name] = value

        super().__init__(display_name=display_name, **other_kwargs)

        for name, comp in registerable.items():
            self._register_component(name, comp)


