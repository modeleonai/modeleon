# SPDX-License-Identifier: Apache-2.0
"""The blend spec — ``mo.blend`` inside the tracks declaration (§16.9).

A doctrine-free SPLICE-AND-CONTINUE rule over declared tracks:

    tracks=mo.Tracks('факт', 'бюджет',
                     blend=mo.blend(given='факт', follow='бюджет',
                                    until='2026-02'))

``given`` names the track whose DATA anchors the synthesized series up
to (and including) the ``until`` period; ``follow`` names the track
that continues after it; ``name`` (default ``'live'``) is the
synthesized track every tracked line grows at materialization.

The synthesized series is a real track in the value layer, so the
broadcast law and the rank-lifting protocol carry it through formulas
UNCHANGED — memoryless lines coincide with the output splice, stateful
lines (recurrence / cumsum / lag) re-anchor because their chains roll
over the operands' live flows (§16.9: one law, three behaviors).

The engine attaches no finance meaning to any of it — «given» is not
«actuals» until a later, explicit role-marker layer says so. ``until``
is a period label of the model grain (the close date under a neutral
name); moving it re-splices every live series on the next run.
"""

from __future__ import annotations

from typing import Optional


class BlendSpec:
    """Validated ``mo.blend(...)`` value — held by the Tracks declaration."""

    __slots__ = ('given', 'follow', 'until', 'name')

    def __init__(self, given: str, follow: str, until: str,
                 name: str = 'live') -> None:
        for arg, val in (('given', given), ('follow', follow),
                         ('until', until), ('name', name)):
            if not isinstance(val, str) or not val.strip():
                raise TypeError(
                    f"mo.blend {arg}= must be a non-empty string; got "
                    f"{val!r}."
                )
        if given == follow:
            raise ValueError(
                "mo.blend given= and follow= name the same track — the "
                "spec splices one track's data INTO another's "
                "continuation; use two different tracks."
            )
        if name in (given, follow):
            raise ValueError(
                f"mo.blend name={name!r} collides with a source track — "
                f"the synthesized series needs its own name."
            )
        self.given = given
        self.follow = follow
        self.until = until
        self.name = name

    def __repr__(self) -> str:
        return (f"mo.blend(given={self.given!r}, follow={self.follow!r}, "
                f"until={self.until!r}, name={self.name!r})")


def blend(given: str, follow: str, until: str, name: str = 'live') -> BlendSpec:
    """Build the splice-and-continue spec for ``mo.Tracks(blend=...)``."""
    return BlendSpec(given, follow, until, name)
