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

import os
import warnings
import weakref
from contextvars import ContextVar
from dataclasses import dataclass, replace
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


def _is_read_only_property(cls: type, name: str) -> bool:
    """Whether ``name`` is a property of ``cls`` that cannot be assigned
    (``code``, ``description``, ``excel_props``, ``python_name``, ``path``,
    ``id`` … on every container, plus any a subclass declares)."""
    attr = getattr(cls, name, None)
    return isinstance(attr, property) and attr.fset is None


def _refuse_read_only_name(cls: type, name: str) -> None:
    """Refuse ``name`` as a component name when the node already answers
    it with a read-only property.

    ``mv.<name>`` would keep returning the property, so the component
    could never be reached by that name — yet it would still be written
    to the workbook. Callers check this before registering anything, so
    a refused name leaves the container exactly as it was.
    """
    if _is_read_only_property(cls, name):
        raise AttributeError(
            f"Cannot name a component `{name}`: `{name}` is a read-only "
            f"property of {cls.__name__}, so `.{name}` could never return "
            f"the component. Pick another name for it."
        )


# ─── Loop containers — rows that read each other across periods ─────
#
# A container built with ``loop=True`` (or a class declared
# ``class X(mo.MultiVariableClass, loop=True)``) lets its rows be read
# before the lines that assign them: a loop that crosses a period
# boundary, like a debt block (see :mod:`modeleon.core.loops`). Reading
# a row it does not have yet hands out a PLACEHOLDER; assigning the row
# turns the placeholder into it. Nothing changes for any other container.

#: Names that are settings of a container, never rows.
_SETTING_SLOTS = frozenset({
    'default_grain', 'default_start', 'default_periods',
    'default_excel_view', 'default_header', 'tracks',
})

#: The engine's own source directory: attribute probes made from here
#: (``getattr(mv, name, None)`` while writing a workbook, …) never
#: create rows in a loop container.
_ENGINE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + os.sep


@dataclass(frozen=True)
class _LoopContext:
    """Where a placeholder finds its horizon when its container has
    none yet (a class instance is built before it joins the model).
    Immutable, in a context variable: threads and tasks never share it."""
    entered: tuple = ()     # containers whose ``with`` block is open
    windowed: Any = None    # weak reference to the model given a horizon last


_loop_context: ContextVar[_LoopContext] = ContextVar(
    'modeleon_loop_context', default=_LoopContext()
)


def is_loop_container(mv: Any) -> bool:
    """True for a container built with ``loop=True`` (or of a loop class)."""
    return bool(mv.__dict__.get('_loop', getattr(type(mv), '_loop', False)))


#: Source directories whose attribute probes never create rows, beside
#: the engine's own — an extension registers its package here.
_TRUSTED_DIRS: list = [_ENGINE_DIR]


def trust_code_in(directory: str) -> None:
    """Treat attribute reads made from code under ``directory`` as the
    engine's own: a probe (``getattr(mv, name, None)``) from there never
    creates a row in a loop container. For extensions that inspect
    models."""
    path = os.path.abspath(directory) + os.sep
    if path not in _TRUSTED_DIRS:
        _TRUSTED_DIRS.append(path)


def is_engine_frame(frame: Any) -> bool:
    """True for a frame running the engine's (or a trusted extension's) code."""
    if frame is None:
        return False
    filename = frame.f_code.co_filename
    return any(filename.startswith(d) for d in _TRUSTED_DIRS)


def _note_window(mv: Any) -> None:
    """Remember the model most recently given a horizon, for a class
    instance built before it joins it (see :func:`_horizon_for`). Weakly:
    remembering must not keep a model alive."""
    from .model import Model
    if isinstance(mv, Model):
        _loop_context.set(replace(_loop_context.get(), windowed=weakref.ref(mv)))


def _horizon_for(mv: Any) -> Optional[int]:
    """The period count a placeholder in ``mv`` holds: ``mv``'s own
    window; else the innermost open ``with`` block's; else the container
    most recently given a horizon — a class instance is built before it
    joins the model, and borrows the model's."""
    from .multi_variable import MultiVariableClass
    from .time import resolve_default_window
    ctx = _loop_context.get()
    for candidate in (mv, *reversed(ctx.entered)):
        window = resolve_default_window(candidate)
        if window is not None and window.periods:
            return int(window.periods)
    # Only a class instance still being built borrows a horizon it was
    # not given — the model built last — and the borrow is checked when
    # the instance joins its model (see :func:`check_borrowed_horizon`).
    # Any other container without one has simply not joined a model.
    model = ctx.windowed() if ctx.windowed is not None else None
    if (model is not None and isinstance(mv, MultiVariableClass)
            and getattr(mv, '_parent', None) is None):
        window = resolve_default_window(model)
        if window is not None and window.periods:
            mv.__dict__['_borrowed_horizon'] = int(window.periods)
            return int(window.periods)
    return None


def ambient_window() -> Any:
    """The window a formula built right now lives in: the innermost open
    ``with`` block's, else the model given a horizon last — the order a
    loop placeholder finds its period count by (:func:`_horizon_for`), so
    ``mo.time`` and a placeholder always agree on the length."""
    from .time import resolve_default_window, section_window_refusal
    ctx = _loop_context.get()
    for candidate in reversed(ctx.entered):
        refusal = section_window_refusal(candidate)
        if refusal is not None:
            raise ValueError(f"mo.time is read where the window is not whole: {refusal}")
        window = resolve_default_window(candidate)
        if window is not None and window.periods:
            return window
    model = ctx.windowed() if ctx.windowed is not None else None
    if model is not None:
        return resolve_default_window(model)
    return None


#: Set the first time ``mo.time`` is read in this process: until then no
#: Variable can read time, and joining a container checks nothing.
_TIME_READ = [False]


def _time_reads(var: Any) -> Any:
    """The ``mo.time`` Variables ``var`` is, or reads through unnamed
    intermediates (a named row it reads was checked when it joined)."""
    from .expr import TimeRef
    seen: set = set()
    stack = [var]
    while stack:
        v = stack.pop()
        if id(v) in seen:
            continue
        seen.add(id(v))
        expr = getattr(v, '_expr', None)
        if isinstance(expr, TimeRef):
            yield v
            continue
        if expr is None:
            continue
        for ref in expr.iter_refs():
            if not getattr(ref, '_reads_time', False):
                continue
            if isinstance(getattr(ref, '_expr', None), TimeRef) or (
                    getattr(ref, '_owner', None) is None
                    and getattr(ref, '_parent', None) is None):
                stack.append(ref)


def _window_node(node: Any) -> Any:
    """The container declaring the window ``node`` lives in."""
    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        if getattr(n, 'default_grain', None) is not None:
            return n
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return None


def _fit_time(name: str, var: Any, window: Any, home: Any) -> None:
    """``var`` (or the unnamed values it is built from) reads ``mo.time``;
    it is going to live in ``window``, declared by ``home``."""
    from .time import time_field_series
    for t in list(_time_reads(var)):
        n = len(t._value) if isinstance(t._value, list) else None
        first_day = getattr(window, 'first_day', None)
        if (t._grain == window.grain and t._start == window.start
                and t._first_day == first_day
                and (not window.periods or n == window.periods)):
            continue
        if t is var and window.periods:
            # The line IS time (``q.days = mo.time.days`` read under another
            # window — code written as flat statements, with no open
            # block): time holds everywhere, so it is simply read again
            # for where it lives.
            field = t._expr.field
            t._value = time_field_series(field, window.start, window.grain,
                                         int(window.periods), first_day)
            t._start, t._grain = window.start, window.grain
            t._first_day = first_day
            continue
        def _from(start: Any, first: Any) -> str:
            return first.isoformat() if first is not None else str(start)
        raise ValueError(
            f"{name!r} reads mo.time.{t._expr.field} for {n} {t._grain} "
            f"period(s) from {_from(t._start, t._first_day)}, but it joins a "
            f"container whose window is {window.periods} {window.grain} "
            f"period(s) from {_from(window.start, first_day)}. "
            f"Read time inside the `with` block of the container the line "
            f"belongs to."
        )
    if home is not None:
        home.__dict__.setdefault('_time_read_by', name)


def check_time_window(parent: Any, name: str, component: Any) -> None:
    """Every line reading ``mo.time`` joins a container of the SAME window.

    Time is read where the formula is built — the open ``with`` block,
    else the model built last. A line that then joins a container of
    another window would carry that other model's dates. A container
    joining (a section, a class instance) brings its lines with it; each
    lives in its own window if the container declares one, else in the
    parent's."""
    if not _TIME_READ[0]:
        return
    from .multi_variable import MultiVariableBase
    from .time import resolve_default_window
    from .variable import Variable
    parent_home = _window_node(parent)
    parent_window = resolve_default_window(parent_home) if parent_home is not None else None
    if isinstance(component, Variable) and not isinstance(component, MultiVariableBase):
        if getattr(component, '_reads_time', False) and parent_window is not None \
                and parent_window.start is not None:
            _fit_time(name, component, parent_window, parent_home)
        return
    for _path, var in component.walk():
        if not isinstance(var, Variable) or isinstance(var, MultiVariableBase):
            continue
        if not getattr(var, '_reads_time', False):
            continue
        home, node, seen = None, getattr(var, '_owner', None), set()
        while node is not None and id(node) not in seen:
            seen.add(id(node))
            if getattr(node, 'default_grain', None) is not None:
                home = node
                break
            if node is component:
                break
            node = getattr(node, '_owner', None) or getattr(node, '_parent', None)
        window = resolve_default_window(home) if home is not None else parent_window
        home = home if home is not None else parent_home
        if window is not None and window.start is not None:
            _fit_time(f"{name}.{getattr(var, '_python_name', None) or 'row'}", var, window, home)


def refuse_window_change(node: Any, setting: str) -> None:
    """A window's lines computed their time when they were built: moving
    the window afterwards would leave them on the old dates."""
    site = node.__dict__.get('_time_read_by')
    if site is not None:
        raise ValueError(
            f"{setting} changes a window whose lines already read mo.time "
            f"(first: {site!r}); they were computed for the old window. Set "
            f"{setting} before the first line that reads time, or build the "
            f"model again."
        )


def check_borrowed_horizon(parent: Any, component: Any) -> None:
    """A loop class instance built outside its model borrowed the horizon
    of the model built last; joining a model with another horizon would
    leave its loop computed over the wrong number of periods."""
    borrowed = component.__dict__.get('_borrowed_horizon')
    if borrowed is None:
        return
    from .time import resolve_default_window
    window = resolve_default_window(parent)
    actual = int(window.periods) if window is not None and window.periods else None
    if actual is not None and actual != borrowed:
        raise ValueError(
            f"{type(component).__name__} was built outside its model and "
            f"computed its loop over {borrowed} period(s), but it joins a "
            f"model of {actual}. Build it inside the model's `with` block."
        )
    del component.__dict__['_borrowed_horizon']


def _forward_placeholder(mv: Any, name: str, site: Any = None) -> Any:
    """The placeholder standing in for ``mv.<name>`` until it is
    assigned — the same object on every read."""
    forward = mv.__dict__.get('_forward')
    if forward is None:
        forward = mv.__dict__['_forward'] = {}
    placeholder = forward.get(name)
    if placeholder is not None:
        return placeholder
    periods = _horizon_for(mv)
    if not periods:
        raise AttributeError(
            f"'{type(mv).__name__}' has no component '{name}'. A loop block "
            f"holds the place of a row read before its line with a series "
            f"as long as the horizon — and this block has none: declare "
            f"default_periods on the model, and build the block inside the "
            f"model's `with` block."
        )
    from .loops import Dependents
    from .variable import Variable
    placeholder = Variable([0.0] * periods)
    placeholder._python_name = name
    placeholder._forward_of = (mv, name)
    placeholder._forward_dependents = Dependents()
    placeholder._forward_site = site
    forward[name] = placeholder
    return placeholder


def _site_text(site: Any) -> str:
    return f", read at {site[0]}:{site[1]}" if site else ""


def check_loop_rows(mv: Any, refuse: bool = True) -> list:
    """Close the reading-ahead of a loop container: a row read before its
    line (a formula holds it) that no line assigned is an error — raised
    with the place it was read, or, without ``refuse``, returned as
    ``(name, site)`` pairs. Placeholders nothing read are dropped."""
    forward = mv.__dict__.get('_forward')
    if not forward:
        return []
    unresolved = [
        (name, placeholder._forward_site)
        for name, placeholder in forward.items()
        if placeholder._forward_dependents
    ]
    for placeholder in forward.values():
        placeholder._forward_abandoned = True
    forward.clear()
    if unresolved and refuse:
        from .errors import ForwardReferenceError
        label = (getattr(mv, '_python_name', None)
                 or getattr(mv, '_display_name', None) or type(mv).__name__)
        where = ", ".join(f"'{name}'{_site_text(site)}" for name, site in unresolved)
        raise ForwardReferenceError(
            f"{where}: read in {label} before {'its' if len(unresolved) == 1 else 'their'} "
            f"line, but never assigned — a misspelt name, or a line that "
            f"did not run (an `if` branch not taken, or an exception caught "
            f"around it)."
        )
    return unresolved


def _refuse_foreign_placeholder(mv: Any, name: str, component: Any) -> None:
    """A placeholder may only become the row it stands for."""
    forward_of = getattr(component, '_forward_of', None)
    if forward_of is None:
        return
    owner, own_name = forward_of
    if owner is mv and own_name == name:
        return
    from .errors import ForwardReferenceError
    raise ForwardReferenceError(
        f"'{name}' is assigned '{own_name}', which is read ahead but not "
        f"assigned yet — assign '{name}' after '{own_name}'s line."
    )


def _refuse_plain_value(mv: Any, name: str, value: Any) -> None:
    """A name a formula read ahead of its line must become a row."""
    forward = mv.__dict__.get('_forward')
    if not forward or name not in forward:
        return
    if not forward[name]._forward_dependents:
        # Only probed; nothing waits on it.
        forward.pop(name)._forward_abandoned = True
        return
    from .errors import ForwardReferenceError
    raise ForwardReferenceError(
        f"'{name}' is read through mo.lag above as a row, but is assigned a "
        f"{type(value).__name__} — wrap it: mo.Variable(...)."
    )


def _transplant(target: Any, component: Any) -> None:
    """Make ``target`` BE ``component`` — same object, new definition —
    so every formula that holds ``target`` now reads the new row. The
    owner links of ``target`` are kept; its values are its own list."""
    keep = {k: target.__dict__[k] for k in
            ('_owner', '_component_name', '_python_name', '_qualified_id')
            if k in target.__dict__}
    state = dict(component.__dict__)
    if isinstance(state.get('_value'), list):
        # Its own list: an alias must not share the source row's values.
        state['_value'] = list(state['_value'])
    target.__dict__.clear()
    target.__dict__.update(state)
    target.__class__ = type(component)
    if keep.get('_owner') is not None:
        target.__dict__.update(keep)


def _become_forward(placeholder: Any, component: Any) -> Any:
    """Turn the placeholder into ``component`` IN PLACE — same object,
    so every formula that already holds it now holds the real row."""
    from .multi_variable import MultiVariableBase
    from .variable import Variable
    from .errors import ForwardReferenceError

    _owner, name = placeholder._forward_of
    if not isinstance(component, Variable) or isinstance(component, MultiVariableBase):
        raise ForwardReferenceError(
            f"'{name}' is read through mo.lag above as a row, but is "
            f"assigned a {type(component).__name__}."
        )
    if component is placeholder:
        raise ForwardReferenceError(
            f"'{name}' is assigned to itself before it has a value."
        )
    if component._forward_of is not None:
        raise ForwardReferenceError(
            f"'{name}' is assigned '{component._forward_of[1]}', which is "
            f"not assigned yet either."
        )
    horizon = (len(placeholder._value)
               if isinstance(placeholder._value, list) else None)
    if not isinstance(component._value, list) and component.var_type != 'list':
        marker = component._value
        if isinstance(marker, str) and marker.startswith('#') and horizon:
            # An Excel error value standing in for the row is that error
            # in every period.
            component._value = [marker] * horizon
            component.var_type = 'list'
        else:
            raise ForwardReferenceError(
                f"'{name}' is read through mo.lag above, but is assigned a "
                f"single value — lag reads a series."
            )
    from .loops import Dependents

    dependents = placeholder._forward_dependents or Dependents()
    _transplant(placeholder, component)
    placeholder._forward_of = None
    placeholder._forward_dependents = dependents
    placeholder._forward_horizon = horizon
    placeholder._forward_site = None
    # The assigned object is now a detached twin of the row; a local
    # name may still hold it, so it settles with the rest.
    if component._awaits:
        dependents.append(component)
    # What the row itself waits on, everything built on it waits on too
    # — those Variables were marked when the row was a bare placeholder
    # that waited on nothing. Let the other placeholders see them.
    others = {p for p in (placeholder._awaits or ()) if p is not placeholder}
    if others:
        chain = [placeholder, *dependents]
        for dependent in dependents:
            if dependent is placeholder:
                continue
            waits = dependent._awaits or set()
            waits.update(others)
            dependent._awaits = waits
        for other in others:
            siblings = other._forward_dependents
            if siblings is not None:
                for dependent in chain:
                    siblings.append(dependent)  # once each: see Dependents
    return placeholder


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
    any root exists; their identity resolves once a root adopts them).
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
        ctx = _loop_context.get()
        _loop_context.set(replace(ctx, entered=ctx.entered + (self,)))
        return self  # type: ignore[return-value]  # mixin self is the concrete MV at runtime

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        ctx = _loop_context.get()
        entered = ctx.entered
        if entered and entered[-1] is self:
            entered = entered[:-1]
            _loop_context.set(replace(ctx, entered=entered))
        # A loop container's (outermost) block ends: a row it read ahead
        # but never assigned is reported here (on a clean exit), or dropped.
        if self.__dict__.get('_forward') and not any(e is self for e in entered):
            check_loop_rows(self, refuse=exc_type is None)
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

        # A name the node answers itself (``mv.code``, ``mv.description``
        # …) is refused before anything below changes state.
        _refuse_read_only_name(type(self), name)

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

            if (is_loop_container(self) and isinstance(existing, Variable)
                    and not isinstance(existing, MultiVariableBase)):
                # The loop was solved with the old row; the rows built on
                # it would keep reading it. A loop is re-built whole.
                label = self._python_name or self._display_name or type(self).__name__
                raise ValueError(
                    f"'{name}' already exists in the loop container {label}: "
                    f"a loop cannot be changed one row at a time. Build the "
                    f"container again (e.g. `m.{label} = mo.MultiVariable("
                    f"..., loop=True)`) and assign its rows anew — or, in a "
                    f"notebook, keep that line in the same notebook cell as the loop."
                )

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

        if isinstance(component, MultiVariableBase):
            check_borrowed_horizon(self, component)
        check_time_window(self, name, component)

        # Clone-on-ownership-change.
        if isinstance(component, Variable) and not isinstance(component, MultiVariableBase):
            if component._owner is not None and component._owner is not self:
                component = component.copy()
        elif isinstance(component, MultiVariableBase):
            if component._parent is not None and component._parent is not self:
                component = component._clone()

        # A row read before this line (a loop broken by ``lag``): the
        # placeholder that took those reads becomes the component, so
        # every formula holding it points at the real row.
        settle_forward = None
        _refuse_foreign_placeholder(self, name, component)
        forward = self.__dict__.get('_forward')
        if forward and name in forward:
            placeholder = forward.pop(name)
            if placeholder._forward_dependents:
                component = _become_forward(placeholder, component)
                settle_forward = component
            else:
                # Only probed (``hasattr`` / ``getattr``), never read by a
                # formula: nothing waits on it.
                placeholder._forward_abandoned = True

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
        # (see ``Variable._on_adopted``). Runs here for the direct
        # adoptee (the ancestor chain is complete the moment
        # back-pointers are set) and again when a floating subtree is
        # mounted under a rooted parent. Idempotent.
        #
        # Walks ``_components`` DIRECTLY — never ``_collect_variables``,
        # which computes ``var.id`` and would freeze FLOATING qpaths onto
        # a subtree that gets its root only later (the projection
        # assembly builds whole trees before any root exists).
        _run_adoption_hooks(component)

        if settle_forward is not None:
            from .loops import settle
            settle(settle_forward)

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
        # A loop container's copy is one too: the flag of a
        # ``MultiVariable(..., loop=True)`` lives on the instance (a loop
        # class carries it on its type, which the copy shares).
        if '_loop' in self.__dict__:
            clone._loop = self._loop
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
