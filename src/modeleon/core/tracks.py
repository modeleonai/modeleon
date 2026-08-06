# SPDX-License-Identifier: Apache-2.0
"""Tracks — the value of a track-axised Variable (§14.2.2, §16).

A tracked value is a mapping **role → track**, where each track is an
ordinary time series (a list over the ambient window) or a scalar:

    revenue._value = TrackValues({'plan': [...24 numbers], 'actual': [...]})

The design's two-tier spine in one type: TIME lives inside each track
(positions under the ambient window, projected by the grain lens);
the FINITE axis is the dict key. Non-axised Variables never see this
type — their ``_value`` stays a flat list or scalar, byte-identical
behavior.

Keys are ROLES (machine keys, §17.2) — plain strings; the closed
vocabulary (actual/plan/forecast) and the role→label display mapping are
the authoring layer's law (``mo.Tracks``, P2), not the value layer's:
this type is the generic mechanism, the doctrine lives at the door.

The broadcast law (§15.3.4) lives in ``combine``:

- identical role sets → element-wise zip per track;
- a plain series / scalar broadcasts into EVERY track;
- differing role sets → a LOUD teaching error, never a silent
  intersection (xarray's inner-join regret; TM1's feeders).
"""

from __future__ import annotations

from typing import Any, Callable, Dict, Iterable, List, Optional, Union

Fiber = Union[List[Any], Any]


class TrackValues:
    """Mapping role → track. Immutable-ish value object for ``_value``."""

    __slots__ = ('_tracks',)

    def __init__(self, tracks: Dict[str, Fiber]) -> None:
        if not isinstance(tracks, dict) or not tracks:
            raise TypeError(
                "TrackValues needs a non-empty {role: series} dict, e.g. "
                "{'plan': [...], 'actual': [...]}."
            )
        lengths = set()
        for role, track in tracks.items():
            if not isinstance(role, str) or not role.strip():
                raise TypeError(
                    f"track keys are role strings; got {role!r}."
                )
            if isinstance(track, dict) or isinstance(track, TrackValues):
                raise TypeError(
                    f"track {role!r} must be a flat series or a scalar — "
                    f"nested tracks would be a second axis, and the axis "
                    f"budget is one (§16.1)."
                )
            if track is None:
                raise TypeError(
                    f"track {role!r} is None — a coordinate must carry a "
                    f"value. (An operand read before its own adoption "
                    f"materializes lands here; attach operands to the "
                    f"model before the variables that read them.)"
                )
            if isinstance(track, list):
                lengths.add(len(track))
        if len(lengths) > 1:
            raise ValueError(
                f"tracks must share one time length; got lengths "
                f"{sorted(lengths)}. Ragged coordinates arrive with the "
                f"role-window law (P2) — until then, pad explicitly."
            )
        self._tracks: Dict[str, Fiber] = dict(tracks)

    # ─── mapping surface ────────────────────────────────────────
    @property
    def roles(self) -> tuple:
        return tuple(self._tracks.keys())

    def __getitem__(self, role: str) -> Fiber:
        return self._tracks[role]

    def __contains__(self, role: str) -> bool:
        return role in self._tracks

    def items(self):
        return self._tracks.items()

    def values(self):
        return self._tracks.values()

    def as_dict(self) -> Dict[str, Fiber]:
        return dict(self._tracks)

    @property
    def time_length(self) -> Optional[int]:
        """Length of the list tracks, or None when all-scalar."""
        for track in self._tracks.values():
            if isinstance(track, list):
                return len(track)
        return None

    def __repr__(self) -> str:
        parts = []
        for role, track in self._tracks.items():
            if isinstance(track, list):
                head = ", ".join(repr(x) for x in track[:3])
                tail = ", …" if len(track) > 3 else ""
                parts.append(f"{role!r}: [{head}{tail}]")
            else:
                parts.append(f"{role!r}: {track!r}")
        return f"TrackValues({{{', '.join(parts)}}})"

    # ─── the broadcast law ──────────────────────────────────────
    @staticmethod
    def combine(left: Any, right: Any,
                op: Callable[[Any, Any], Any]) -> 'TrackValues':
        """Combine two operands, at least one tracked, per §15.3.4.

        ``op`` is the plain element-wise combiner the chokepoint already
        uses for flat values (it handles scalars, lists, error markers).
        """
        lf = isinstance(left, TrackValues)
        rf = isinstance(right, TrackValues)
        if lf and rf:
            if set(left.roles) != set(right.roles):
                raise ValueError(
                    f"coordinate sets differ: {sorted(left.roles)} vs "
                    f"{sorted(right.roles)} — arithmetic never silently "
                    f"intersects coordinates. Slice one side "
                    f"(.at(track='...')) or give both the same "
                    f"coordinates."
                )
            return TrackValues({
                role: op(left[role], right[role]) for role in left.roles
            })
        if lf:
            return TrackValues({
                role: op(track, right) for role, track in left.items()
            })
        return TrackValues({
            role: op(left, track) for role, track in right.items()
        })

    @staticmethod
    def lift(fn: Callable[..., Any], *operands: Any) -> 'TrackValues':
        """Rank-lifting (§15.3.5): map ``fn`` over roles.

        Every track-carrying operand must share one role set (the mismatch law
        again); plain operands are passed through unchanged to every
        role's call. This one wrapper is how lag / recurrence chains /
        cumsum / aggregates / IF operate per coordinate — no per-
        function whitelist to drift.
        """
        role_sets = [set(o.roles) for o in operands if isinstance(o, TrackValues)]
        if not role_sets:
            raise TypeError("lift() needs at least one track-carrying operand")
        first = role_sets[0]
        if any(rs != first for rs in role_sets[1:]):
            raise ValueError(
                f"coordinate sets differ across operands: "
                f"{[sorted(rs) for rs in role_sets]} — slice or align "
                f"before the operation."
            )
        roles = next(o for o in operands if isinstance(o, TrackValues)).roles
        out: Dict[str, Fiber] = {}
        for role in roles:
            args = [
                o[role] if isinstance(o, TrackValues) else o for o in operands
            ]
            out[role] = fn(*args)
        return TrackValues(out)
