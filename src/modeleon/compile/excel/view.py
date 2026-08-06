# SPDX-License-Identifier: Apache-2.0
"""ExcelView — a declared-once, cascading presentation view, *as a model*.

``ExcelView`` is a :class:`MultiVariable`, organised into **sub-views** (groups
of related settings, each a sub-model):

    orient                          geometry  — period axis direction
    font       {name, size}         typography
    bands      {section, …}         structure — section/band styling
    cell_types {on, colors{…}}      cell-colour by kind (input/reference)
    timeline   {header, label_format, rows[…]}   the period-label header

So a view is a real tree you navigate (``view.timeline.label_format``,
``view.cell_types.colors.input``). Renderers don't read the tree — they read a
**flat** :class:`ResolvedView` snapshot produced by :func:`resolve_excel_view`,
which walks to the nearest ``default_excel_view`` and fills unset leaves from the
global house default ``mo.default_excel_view``. Authoring/display is the grouped
model; consumption is the flat snapshot. Fields are PRESENTATION only — a view
never owns *what periods exist* (that's the model's timeline / projection;
ADR-010).
"""

from __future__ import annotations

from dataclasses import dataclass
from dataclasses import fields as _dc_fields
from html import escape
from typing import Any, Dict, Optional

from ...core.multi_variable import MultiVariable, MultiVariableBase
from ...core.variable import Variable

#: The authoring groups — the sub-views of an ExcelView.
_GROUPS = ("orient", "font", "bands", "cell_types", "timeline",
           "tracks", "grain")


@dataclass(frozen=True)
class ResolvedView:
    """A compiled, plain-value snapshot — the FLAT fields the renderers read.

    Groups are flattened (``timeline.header`` → ``time_header``, etc.) so the
    field-level merge over the house default inherits each leaf independently."""

    orient: Optional[str] = None
    #: How tracked lines print (§16 P3): 'blend' — the display-default
    #: (live) row only; 'rows' — default row + labeled per-track rows;
    #: 'compare' — period-major column groups (later work).
    tracks: Optional[str] = None
    #: The view's grain lens — which grain this view prints/shows at.
    #: None = the model's native grain.
    grain: Optional[str] = None
    base_font: Optional[Dict[str, Any]] = None
    band_styles: Optional[Dict[str, Dict[str, Any]]] = None
    format_by_type: Any = None
    time_header: Optional[bool] = None
    time_label_format: Optional[str] = None
    type_colors: Optional[Dict[str, str]] = None


_RESOLVED_FIELDS = tuple(f.name for f in _dc_fields(ResolvedView))


def _to_node(name: str, value: Any) -> Any:
    """A config value as a tree node: an already-built node as-is, a sub-model
    for a dict (authoring sugar), else a ``Variable``."""
    if isinstance(value, (Variable, MultiVariableBase)):
        return value
    if isinstance(value, dict):
        mv = MultiVariable(name)
        for k, v in value.items():
            setattr(mv, str(k), _to_node(str(k), v))
        return mv
    return Variable(value, display_name=name)


def _node_value(node: Any) -> Any:
    """A tree node back to a plain value: a dict for a sub-model, else the value."""
    if isinstance(node, Variable):
        return node._value
    if isinstance(node, MultiVariableBase):
        return {k: _node_value(c) for k, c in node._components.items()}
    return node


class ExcelView(MultiVariable):
    """Presentation settings for the Excel + notebook renderers, as a model.

    Settings are grouped into sub-views (``font`` / ``bands`` / ``cell_types`` /
    ``timeline``) plus the ``orient`` scalar; each set group is a child sub-model
    (dict literals are authoring sugar — they become sub-models). Read by
    renderers through :func:`resolve_excel_view`, which flattens the groups.
    """

    _TRACK_MODES = ("blend", "rows", "compare")

    def __init__(
        self,
        orient: Optional[str] = None,
        font: Any = None,
        bands: Any = None,
        cell_types: Any = None,
        timeline: Any = None,
        tracks: Optional[str] = None,
        grain: Optional[str] = None,
    ) -> None:
        super().__init__("ExcelView")
        if tracks is not None and tracks not in self._TRACK_MODES:
            raise ValueError(
                f"ExcelView tracks= must be one of "
                f"{', '.join(self._TRACK_MODES)}; got {tracks!r}."
            )
        if grain is not None:
            from ...core.time import GRAINS
            if grain not in GRAINS:
                raise ValueError(
                    f"ExcelView grain= must be one of {', '.join(GRAINS)}; "
                    f"got {grain!r}."
                )
        for name, value in (
            ("orient", orient),
            ("font", font),
            ("bands", bands),
            ("cell_types", cell_types),
            ("timeline", timeline),
            ("tracks", tracks),
            ("grain", grain),
        ):
            if value is not None:
                setattr(self, name, _to_node(name, value))

    def _grouped(self) -> Dict[str, Any]:
        return {f: _node_value(self._components.get(f)) for f in _GROUPS}

    def snapshot(self) -> ResolvedView:
        """Compile this grouped view into the flat :class:`ResolvedView`."""
        g = self._grouped()
        timeline = g.get("timeline") or {}
        cell_types = g.get("cell_types") or {}
        return ResolvedView(
            orient=g.get("orient"),
            tracks=g.get("tracks"),
            grain=g.get("grain"),
            base_font=g.get("font"),
            band_styles=g.get("bands"),
            format_by_type=cell_types.get("on"),
            type_colors=cell_types.get("colors"),
            time_header=timeline.get("header"),
            time_label_format=timeline.get("label_format"),
        )

    def _repr_html_(self) -> str:
        """A flat one-panel summary of the effective settings (no per-group tabs)."""
        s = self.snapshot()

        def _font(bf: Optional[Dict[str, Any]]) -> str:
            if not bf:
                return "—"
            parts: list[str] = [str(bf["name"])] if bf.get("name") else []
            if bf.get("size") is not None:
                parts.append(str(bf["size"]))
            return " · ".join(parts) or "—"

        def _cell_types(on: Any, colors: Optional[Dict[str, str]]) -> str:
            state = "on" if (on is True or isinstance(on, dict)) else (
                "off" if on is False else None
            )
            palette = colors if isinstance(colors, dict) else (
                on if isinstance(on, dict) else None
            )
            parts: list[str] = [state] if state else []
            if isinstance(palette, dict):
                parts += [f"{k} {v}" for k, v in palette.items()]
            return " · ".join(parts) or "—"

        def _timeline(header: Any, label_format: Optional[str]) -> str:
            parts: list[str] = []
            if header:
                parts.append("header")
            if label_format:
                parts.append(str(label_format))
            return " · ".join(parts) or "—"

        rows = (
            ("orient", s.orient or "—"),
            ("tracks", s.tracks or "—"),
            ("grain", s.grain or "—"),
            ("font", _font(s.base_font)),
            ("cell_types", _cell_types(s.format_by_type, s.type_colors)),
            ("timeline", _timeline(s.time_header, s.time_label_format)),
            ("bands", ", ".join(s.band_styles) if s.band_styles else "—"),
        )
        lbl = (
            "padding:5px 12px;background:#f7f8f7;border-bottom:1px solid #e2e8e4;"
            "border-right:1px solid #e2e8e4;text-align:left;font-weight:600;color:#262c38;"
        )
        val = (
            "padding:5px 12px;border-bottom:1px solid #e2e8e4;text-align:left;"
            "font-family:'JetBrains Mono',ui-monospace,monospace;font-size:12px;color:#4a5060;"
        )
        body = "".join(
            f'<tr><td style="{lbl}">{k}</td>'
            f'<td style="{val}">{escape(str(v))}</td></tr>'
            for k, v in rows
        )
        return (
            "<table style=\"border-collapse:collapse;"
            "font-family:'Inter',-apple-system,system-ui,sans-serif;font-size:13px;"
            "color:#262c38;background:#ffffff;border:1px solid #c4cec6;border-radius:4px;"
            "overflow:hidden;\">"
            "<tr><td colspan=\"2\" style=\"padding:6px 12px;background:#07464a;"
            "color:#ffffff;font-weight:600;letter-spacing:0.02em;\">ExcelView</td></tr>"
            f"{body}</table>"
        )

    def extend(self, **overrides: Any) -> "ExcelView":
        """Return a copy with ``overrides`` (whole sub-views) applied."""
        unknown = set(overrides) - set(_GROUPS)
        if unknown:
            raise ValueError(
                f"ExcelView.extend: unknown group(s) {sorted(unknown)}; "
                f"known groups are {sorted(_GROUPS)}"
            )
        merged = self._grouped()
        merged.update(overrides)
        return ExcelView(**{k: v for k, v in merged.items() if v is not None})

    @classmethod
    def engine_default(cls) -> "ExcelView":
        """The zero-config floor — no groups set."""
        return cls()

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, ExcelView):
            return NotImplemented
        return self.snapshot() == other.snapshot()

    __hash__ = object.__hash__


def _global_default() -> Optional[ExcelView]:
    """The process-wide house default — ``modeleon.default_excel_view``, read
    live so a global override takes effect. ``None`` before the package finishes
    importing or if the attribute has been cleared."""
    try:
        import modeleon
        g = getattr(modeleon, "default_excel_view", None)
        if isinstance(g, ExcelView):
            return g
    except Exception:
        pass
    return None


def _merge(snap: ResolvedView, default: ResolvedView) -> ResolvedView:
    """Fill ``snap``'s ``None`` leaves from ``default`` — flat, so a partial
    sub-view (say only ``timeline.header``) still inherits ``label_format`` etc."""
    return ResolvedView(
        **{
            f: (getattr(snap, f) if getattr(snap, f) is not None
                else getattr(default, f))
            for f in _RESOLVED_FIELDS
        }
    )


def resolve_excel_view(node: Any) -> ResolvedView:
    """The effective :class:`ResolvedView` for ``node`` — resolved HIERARCHICALLY.

    Walks ``_owner`` → ``_parent`` from ``node`` up to the root, collecting EVERY
    ``default_excel_view`` pointer on the way (not just the nearest). Fields
    cascade like CSS: the *closest* level that sets a field wins, falling back up
    the tree, with the process-wide ``mo.default_excel_view`` as the top-level
    default (the floor). With no pointer anywhere, that global default is
    returned.
    """
    seen: set[int] = set()
    n = node
    chain: list[ResolvedView] = []  # nearest-first
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        v = getattr(n, "default_excel_view", None)
        if isinstance(v, ExcelView):
            chain.append(v.snapshot())
        n = getattr(n, "_owner", None) or getattr(n, "_parent", None)
    g = _global_default()
    result = g.snapshot() if g is not None else ResolvedView()
    # Fold top-down (root-most first) so each lower level overrides its ancestors
    # and the nearest level wins; the global default stays the floor.
    for snap in reversed(chain):
        result = _merge(snap, result)
    return result
