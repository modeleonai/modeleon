# SPDX-License-Identifier: Apache-2.0
"""The tracks declaration — ``mo.Tracks`` on the Model.

The axis is declared ONCE per model, naming its tracks, and inherited
ambiently — exactly like the time window:

    company = mo.Model('company',
        tracks=mo.Tracks('actual', 'budget'),
        default_grain='month', default_start='2025-01', default_periods=24,
    )

Track names are USER CONTENT — any words, any language, any count —
the same way variable names are (``mo.Tracks('факт', 'бюджет')`` works
the same). A keyword spelling attaches a display label distinct from
the name (``mo.Tracks(budget='Budget 2026')``); the positional spelling
uses the name as its own label.

The engine attaches NO semantics to any track name: tracks are parallel
series inside one Variable — broadcasting arithmetic, slicing, lifts —
pure mechanism. Even ``blend=`` only splices the tracks it is told to
name. Financial role BEHAVIORS (which track holds actuals, what a
scenario may touch, variance sign) are not the engine's business: code
built on top that needs them assigns a role to a track EXPLICITLY
(e.g. by marking "this track plays the actuals role"). Declaring tracks
never requires declaring roles.
"""

from __future__ import annotations

from typing import Dict, Tuple


class Tracks:
    """Name → display-label declaration of the model's tracks axis."""

    __slots__ = ('_labels', 'blend')

    def __init__(self, *names: str, blend=None, **labeled: str) -> None:
        labels: Dict[str, str] = {}
        for n in names:
            if not isinstance(n, str) or not n.strip():
                raise TypeError(
                    f"track names are non-empty strings; got {n!r}."
                )
            labels[n] = n
        for n, lb in labeled.items():
            if not isinstance(lb, str) or not lb.strip():
                raise TypeError(
                    f"the label for track {n!r} must be a non-empty "
                    f"string (it is the display word; omit it to use "
                    f"the name itself)."
                )
            labels[n] = lb
        if not labels:
            raise TypeError(
                "mo.Tracks() declares the model's tracks — name them: "
                "mo.Tracks('actual', 'budget') or "
                "mo.Tracks(budget='Budget 2026', actual='Actual')."
            )
        if blend is not None:
            from .blend import BlendSpec
            if not isinstance(blend, BlendSpec):
                raise TypeError(
                    "blend= takes mo.blend(given=..., follow=..., "
                    f"until=...); got {type(blend).__name__}."
                )
            for side in (blend.given, blend.follow):
                if side not in labels:
                    raise ValueError(
                        f"mo.blend names track {side!r}, which is not "
                        f"declared ({', '.join(labels)}) — declare it or "
                        f"fix the name."
                    )
            if blend.name in labels:
                raise ValueError(
                    f"mo.blend name={blend.name!r} collides with a "
                    f"declared track — the synthesized series needs its "
                    f"own name."
                )
        self.blend = blend
        self._labels = labels

    @property
    def names(self) -> Tuple[str, ...]:
        """Declared track names, in declaration order."""
        return tuple(self._labels)

    def label(self, name: str) -> str:
        return self._labels[name]

    def __contains__(self, name: str) -> bool:
        return name in self._labels

    def __repr__(self) -> str:
        inner = ", ".join(
            (repr(n) if n == lb else f"{n}={lb!r}")
            for n, lb in self._labels.items()
        )
        if self.blend is not None:
            inner += f", blend={self.blend!r}"
        return f"mo.Tracks({inner})"


def resolve_tracks_decl(node) -> Tracks | None:
    """The nearest ancestor's ``tracks`` declaration, or None.

    Same ``_owner`` → ``_parent`` walk as the ambient time window —
    the tracks axis is a model-level ambient. It is a Model kwarg rather
    than a Variable, so it never becomes a node of the dependency graph,
    and every descendant sees the same declaration whichever subtree you
    emit.
    """
    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        decl = getattr(n, 'tracks', None)
        if isinstance(decl, Tracks):
            return decl
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return None
