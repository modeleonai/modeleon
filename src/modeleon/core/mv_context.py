# SPDX-License-Identifier: Apache-2.0
"""Lifecycle + component adoption for :class:`MultiVariableBase`.

Owns three things:

- ``__enter__`` / ``__exit__`` — the context-manager protocol.
  ``with mv:`` returns ``self`` and exits without side effects.
  ``with mv as alias:`` is plain Python — the alias is the same object
  as ``mv``, just under a shorter local name. **The ``with`` statement
  itself does not attach anything to ``mv``.** Attachment happens via
  attribute assignment (``mv.child = thing``).
- ``_register_component`` — the adoption logic invoked from
  ``MultiVariableBase.__setattr__`` when a user assigns
  ``mv.attr = some_variable_or_mv``. Wires back-pointers, sets
  ``python_name`` / ``_component_name``, clones the new component if
  it's already owned by a different parent, and crystallizes paths
  when the parent is rooted.
- ``_clone`` — re-parent clone used by ``_register_component`` when a
  component arrives already owned (shared sub-MVs mounted on two
  parents).

The mixin :class:`_MVLifecycle` is inherited by
:class:`MultiVariableBase`. Methods rely on ``self._components``,
``self._component_order``, ``self._parent``, etc. being present —
they're MV internals, not pure functions.

Lazy imports for :class:`Variable` and :class:`MultiVariableBase`
inside methods avoid the circular import that would arise if we
resolved them at module scope.
"""

from __future__ import annotations

import warnings
from typing import Any, Optional, TYPE_CHECKING

from .qpath import QPath

if TYPE_CHECKING:
    from .multi_variable import MultiVariableBase


class ModelStructureWarning(UserWarning):
    """Warning emitted when an MV-tree construction looks suspicious —
    currently only fires when a component slot is reassigned, orphaning
    the previous component. Future diagnostics may be added here.

    Configured below to always fire — Python's default warning dedup
    ('default' filter) suppresses subsequent occurrences at the same
    source line, which hides the issue when a notebook user re-runs
    the cell.
    """


warnings.filterwarnings('always', category=ModelStructureWarning)


def _crystallize_subtree(item: Any, path: QPath) -> None:
    """Recursively set ``_qualified_id`` on an adopted subtree.

    Called by ``_register_component`` when the parent already has a
    non-FLOATING path. Variables are leaves; MVs recurse into their
    components. No-op when adoption happens before the parent itself is
    rooted — paths stay FLOATING and crystallize later, at the moment
    the root adopts its subtree.
    """
    from .multi_variable import MultiVariableBase

    item._qualified_id = path
    if isinstance(item, MultiVariableBase):
        for comp_name in item._component_order:
            child = item._components[comp_name]
            _crystallize_subtree(child, path.child(comp_name))


def _run_adoption_hooks(item: Any) -> None:
    """Run ``Variable._on_adopted`` over ``item`` (a Variable or an MV
    subtree), walking ``_components`` directly.

    Deliberately NEVER touches ``var.id`` / ``_collect_variables`` —
    identity resolution on a still-floating subtree would CACHE the
    floating qpaths (the projection assembly builds whole trees before
    any root exists, and the run resolves their identity later).
    """
    from .multi_variable import MultiVariableBase
    from .variable import Variable

    if isinstance(item, MultiVariableBase):
        for comp_name in item._component_order:
            _run_adoption_hooks(item._components[comp_name])
    elif isinstance(item, Variable):
        item._on_adopted()


def _invalidate_subtree_paths(item: Any) -> None:
    """Recursively clear ``_qualified_id`` so paths re-resolve on next read.

    Used when an adoptee is re-rooted under a new parent — its old
    cached path is stale, and the parent itself may not yet be rooted
    (so :func:`_crystallize_subtree` has nothing to write). Clearing
    lets the lazy ``path`` property walk the owner chain on access.
    """
    from .multi_variable import MultiVariableBase

    item._qualified_id = None
    if isinstance(item, MultiVariableBase):
        for comp_name in item._component_order:
            child = item._components[comp_name]
            _invalidate_subtree_paths(child)


class _MVLifecycle:
    """Context-manager + adoption mixin for :class:`MultiVariableBase`.

    Attributes listed below are provided by the concrete
    :class:`MultiVariableBase` subclass — declared here for
    type-checkers.
    """

    # Attributes provided by MultiVariableBase.
    _components: dict
    _component_order: list
    _display_name: Optional[str]
    _python_name: Optional[str]
    _qualified_id: Any
    _role: str

    # ─── Context-manager protocol ───────────────────────────────

    def __enter__(self) -> "MultiVariableBase":
        return self  # type: ignore[return-value]  # mixin self is the concrete MV at runtime

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        return None

    # ─── Component adoption & cloning ───────────────────────────

    def _register_component(self, name: str, component: Any) -> None:
        """Register a Variable or nested MV as a named component of self.

        Called by ``MultiVariableBase.__setattr__`` whenever a user
        assigns ``mv.attr = thing`` and ``thing`` is a Variable or MV.

        Clone-on-ownership-change: a Variable already owned by another
        MV gets ``.copy()``'d; an MV already parented elsewhere gets
        ``._clone()``'d. This prevents the same object from ending up
        as a live child of two different parents.

        Duplicate-name handling: re-assigning the same name behaves
        like plain Python attribute assignment — the new value wins. A
        ``ModelStructureWarning`` fires so the overwrite is visible
        (important for cell re-runs in notebooks). Idempotent
        self-assignment (``m.x = m.x``) is silent.

        The replaced component is orphaned: its parent / owner links are
        cleared and its qualified path reverts to the floating namespace.

        Model demotion: when a :class:`~modeleon.core.model.Model` is
        adopted into another container it stops being a root, so its
        runtime class is rewritten to :class:`MultiVariable`. The
        memory layout is identical (Model adds no slots or fields
        beyond MultiVariable's), the user's reference still points to
        the same object, and any Model-specific type checks downstream
        correctly report the demoted instance.
        """
        from .model import Model
        from .multi_variable import MultiVariable, MultiVariableBase
        from .variable import Variable

        # Demote a Model to a plain MultiVariable on adoption — Models
        # are top-level roots by contract; once adopted they are sub-
        # trees. The class swap is layout-safe because Model adds no
        # state beyond what MultiVariable already has. Sheet-vs-section
        # rendering is decided positionally by the renderer (1st-level
        # children of the render root become sheets, deeper become
        # sections), so demotion doesn't touch ``_is_sheet`` /
        # ``excel_props`` — the user's explicit markers are respected
        # if they set any.
        if isinstance(component, Model):
            component.__class__ = MultiVariable  # type: ignore[assignment]

        if name in self._components:
            existing = self._components[name]
            if existing is component:
                return  # idempotent self-assignment

            warnings.warn(
                f"Component `{name}` on `{type(self).__name__}"
                f"({self._display_name!r})` was replaced. "
                f"The previous component is orphaned; formulas that "
                f"still reference it will keep rendering against the "
                f"old object.",
                category=ModelStructureWarning,
                stacklevel=3,
            )
            # Detach the old component so its state doesn't leak back.
            self._detach_component(name)

        # Clone-on-ownership-change.
        if isinstance(component, Variable) and not isinstance(component, MultiVariableBase):
            if component._owner is not None and component._owner is not self:
                component = component.copy()
        elif isinstance(component, MultiVariableBase):
            if component._parent is not None and component._parent is not self:
                component = component._clone()

        # Component storage: live object in ``__dict__``; name tracked
        # in ``_component_names`` (the disambiguator vs private attrs).
        self.__dict__[name] = component
        if name not in self._component_names:
            self._component_names.append(name)

        # ``self`` is always a :class:`MultiVariableBase` at runtime — the mixin
        # is only mounted on MultiVariableBase. Cast for the type-checker.
        self_mv: "MultiVariableBase" = self  # type: ignore[assignment]
        if isinstance(component, MultiVariableBase):
            component._parent = self_mv
            component._name_in_parent = name
            component._python_name = name
        elif isinstance(component, Variable):
            component._owner = self_mv
            component._component_name = name
            component._python_name = name

        # The adoptee may have a cached ``_qualified_id`` from a prior
        # life — e.g. a demoted Model whose path was crystallized at
        # ``ROOT.child(name)`` before adoption. Clear the subtree so
        # the lazy ``path`` property re-resolves through the new owner.
        _invalidate_subtree_paths(component)

        # Crystallize qualified path. When the parent is already
        # rooted, descend into the adoptee and set every descendant's
        # ``_qualified_id``. When the parent itself is still floating
        # (typical mid-construction state), skip — crystallization will
        # happen later when an ancestor adopts us with a rooted path.
        if self._qualified_id is not None:
            _crystallize_subtree(component, self._qualified_id.child(name))

        # Adoption hook — schedule materialization + the extent law
        # (§16.6). Runs here for the direct adoptee (the ancestor chain
        # is complete the moment back-pointers are set) and again when a
        # floating subtree is mounted under a rooted parent. Idempotent.
        #
        # Walks ``_components`` DIRECTLY — never ``_collect_variables``,
        # which computes ``var.id`` and would freeze FLOATING qpaths onto
        # a subtree that gets its root only later (the projection
        # assembly builds whole trees before any root exists).
        _run_adoption_hooks(component)

    def _detach_component(self, name: str) -> Any:
        """Orphan and unregister the child named ``name``.

        Clears the child's parent / owner back-links and floats its
        path, then drops it from ``__dict__`` + ``_component_names``.
        Returns the detached component, or ``None`` if there was no such
        component. Shared by the replace branch of
        :meth:`_register_component` and the public
        :meth:`~modeleon.core.multi_variable.MultiVariableBase.remove`.
        """
        from .multi_variable import MultiVariableBase
        from .variable import Variable

        if name not in self._components:
            return None
        existing = self._components[name]
        if isinstance(existing, MultiVariableBase):
            existing._parent = None
            existing._name_in_parent = None
        elif isinstance(existing, Variable):
            existing._owner = None
            existing._component_name = None
        existing._qualified_id = None
        existing._python_name = None
        self.__dict__.pop(name, None)
        if name in self._component_names:
            self._component_names.remove(name)
        return existing

    def _clone(self) -> "MultiVariableBase":
        """Shallow clone: new MV with copied Variables and cloned child MVs.

        Triggered when a parent re-adopts an MV already owned by another
        parent (e.g. a shared assumptions block mounted on two sheets).
        Each copied Variable's ``_expr`` is a ``MethodCall(VarRef(source),
        'copy')`` node — rendered as a cross-sheet reference in Excel.
        """
        from .multi_variable import MultiVariableBase
        from .variable import Variable

        # ``object.__new__(type(self))`` returns a fresh instance of the
        # concrete MV subclass — typed as the same class as ``self``.
        clone: "MultiVariableBase" = object.__new__(type(self))  # type: ignore[assignment]
        MultiVariableBase.__init__(
            clone,
            display_name=self._display_name,
        )
        clone._role = self._role
        clone._python_name = None
        for comp_name in self._component_order:
            comp = self._components[comp_name]
            if isinstance(comp, Variable) and not isinstance(comp, MultiVariableBase):
                clone._register_component(comp_name, comp.copy())
            elif isinstance(comp, MultiVariableBase):
                clone._register_component(comp_name, comp._clone())
            else:
                clone._components[comp_name] = comp
                clone._component_order.append(comp_name)
        return clone
