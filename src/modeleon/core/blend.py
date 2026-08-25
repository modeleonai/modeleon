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

``until`` is OPTIONAL. Omitted, the splice has no boundary at all:
``given`` wins wherever it carries a value and ``follow`` fills every
cell it leaves empty. That is the honest shape when the given track is
typed as events arrive — irregular, incomplete, out of order — instead
of being closed period by period. With a boundary, the only difference
is AFTER it, where the preference flips to ``follow``; before it the
rule is already «given where present, follow otherwise», so a
boundary-less spec is not a weaker guarantee — it is the same rule
without a date to maintain.

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

    def __init__(self, given: str, follow: str, until: Optional[str] = None,
                 name: str = 'live') -> None:
        for arg, val in (('given', given), ('follow', follow),
                         ('name', name)):
            if not isinstance(val, str) or not val.strip():
                raise TypeError(
                    f"mo.blend {arg}= must be a non-empty string; got "
                    f"{val!r}."
                )
        if until is not None and (
            not isinstance(until, str) or not until.strip()
        ):
            raise TypeError(
                f"mo.blend until= must be a period label like '2026-03', "
                f"or omitted for a boundary-less splice; got {until!r}."
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
        until = "" if self.until is None else f"until={self.until!r}, "
        return (f"mo.blend(given={self.given!r}, follow={self.follow!r}, "
                f"{until}name={self.name!r})")


def blend(given: str, follow: str, until: Optional[str] = None,
          name: str = 'live') -> BlendSpec:
    """Build the splice-and-continue spec for ``mo.Tracks(blend=...)``.

    Omit ``until`` for the BOUNDARY-LESS form: ``given`` wins in every
    period where it carries a value, ``follow`` fills the rest. Use it
    when the given track arrives irregularly — a register typed cell by
    cell as events land, rather than closed period by period.
    """
    return BlendSpec(given, follow, until, name)
