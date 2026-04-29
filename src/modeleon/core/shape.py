# SPDX-License-Identifier: Apache-2.0
"""Shape — what dimensions a Variable extends through.

Every Variable has exactly one :class:`Shape`. A scalar Variable's
shape is empty (no axis, no behavior, no keys). A list-valued
Variable's shape names the axis it lives on. A keyed (dict-sourced)
Variable's shape names the keys.

Shape is intentionally small: layout reads ``length`` to allocate
columns; arithmetic reads ``axis`` to check broadcast compatibility;
aggregation reads ``behavior`` to pick the right rule. New dimensions
extend Shape itself rather than adding parallel attributes on
Variable.

The ``axis_owner`` is a *reference* to the live axis-providing
object, not a snapshot. Mutating the owner (replacing its
``axis`` field with a new frozen :class:`Axis`) is observed by every
Variable bound to it on next read — the Anaplan / TM1 pattern of
"adjust the dimension, every line item dimensioned by it follows."
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any, Literal, Optional, Protocol, runtime_checkable
from weakref import WeakSet

from .axis import Axis, PositionalAxis

if TYPE_CHECKING:
    from .variable import Variable


TimeBehavior = Literal["flow", "stock", "point", "rate", "scalar"]


@runtime_checkable
class AxisProvider(Protocol):
    """Anything that owns a live :class:`Axis` and notifies dependents on change.

    Concrete providers expose the current axis as a property and
    invalidate downstream caches when their state mutates.
    """

    @property
    def axis(self) -> Axis: ...

    def _invalidate_dependents(self) -> None: ...


@dataclass(frozen=True, slots=True)
class Shape:
    """Dimensions of a Variable.

    Fields are independent and all optional:

    - ``axis_owner``: a reference to an :class:`AxisProvider`. ``None``
      means the Variable is scalar along this dimension.
    - ``behavior``: how this Variable aggregates across periods —
      ``"flow"`` (sums), ``"stock"`` (last value), ``"rate"``
      (weighted-mean), ``"point"`` (broadcast). ``None`` defers the
      decision to the call site.
    - ``keys``: named-list dimension orthogonal to the axis (e.g.
      ``("US", "EU", "APAC")``). Used by dict-sourced Variables.

    The ``axis`` and ``length`` properties dereference ``axis_owner``
    on every read so live binding works without callers having to
    re-fetch.
    """

    axis_owner: Optional[AxisProvider] = None
    behavior: Optional[TimeBehavior] = None
    keys: Optional[tuple[str, ...]] = None

    @property
    def axis(self) -> Optional[Axis]:
        """Current :class:`Axis` from the bound owner, or ``None`` for scalars."""
        return self.axis_owner.axis if self.axis_owner is not None else None

    @property
    def length(self) -> int:
        """Cell count along this Shape.

        Resolution order: axis length (if axis-bound), keys count
        (if keyed-only), 1 (scalar). When both axis and keys are
        present the cross-product is not yet expressed by ``length``
        — current design treats them as orthogonal one-axis-at-a-time.
        """
        ax = self.axis
        if ax is not None:
            return ax.length
        if self.keys is not None:
            return len(self.keys)
        return 1

    def is_scalar(self) -> bool:
        """True when no axis and no keys — a single value."""
        return self.axis_owner is None and self.keys is None


@dataclass(frozen=True, slots=True)
class _AnonymousPositionalProvider:
    """Tiny :class:`AxisProvider` wrapping a frozen :class:`PositionalAxis`.

    Used when a Variable is constructed with a list and no explicit
    axis — the value's length implies a positional axis but there is
    no named owner. The provider satisfies the :class:`AxisProvider`
    Protocol with a no-op invalidation hook (the wrapped axis is
    immutable; nothing to invalidate).
    """

    axis: PositionalAxis

    def _invalidate_dependents(self) -> None:
        """No-op — the wrapped axis is frozen and shared only by reference."""
        return None


class MutableAxisProvider:
    """Base for :class:`AxisProvider` implementations whose axis can change.

    Holds the current frozen :class:`Axis` and a weak set of bound
    Variables. ``adjust(**kw)`` atomically replaces the axis with a
    new frozen instance and notifies every dependent — Variables
    bound by reference observe the change without needing to re-read.

    Concrete subclasses can override :meth:`adjust` to validate
    their domain-specific kwargs (e.g. fiscal-year offsets, calendar
    frequencies) before delegating here.
    """

    def __init__(self, axis: Axis) -> None:
        self._axis = axis
        # Weak set so dropped Variables don't keep this provider alive
        # past their own lifetime.
        self._dependents: "WeakSet[Variable]" = WeakSet()

    @property
    def axis(self) -> Axis:
        """The current frozen axis. Replaced atomically by :meth:`adjust`."""
        return self._axis

    def _register_dependent(self, var: "Variable") -> None:
        """Attach a Variable to this provider so it receives invalidations.

        Called automatically when a Variable is constructed with
        ``axis=`` pointing at a :class:`MutableAxisProvider`. The
        weak set lets the Variable be garbage-collected normally.
        """
        self._dependents.add(var)

    def _invalidate_dependents(self) -> None:
        """Notify every bound Variable that the axis has changed.

        Each dependent's ``_on_axis_invalidated`` is called once so
        the Variable can clear any caches it derived from the prior
        axis state.
        """
        for var in list(self._dependents):
            var._on_axis_invalidated()

    def adjust(self, **kw: Any) -> None:
        """Replace the current axis with a new frozen instance and
        invalidate dependents.

        Accepted kwargs are whatever :func:`dataclasses.replace`
        accepts on the wrapped axis class — ``length``, ``name``, and
        any subclass-specific fields. The replacement is atomic from
        the caller's perspective.
        """
        self._axis = replace(self._axis, **kw)
        self._invalidate_dependents()
