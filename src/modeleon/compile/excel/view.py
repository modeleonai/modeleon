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
           "tracks", "grain", "formats", "nesting", "meta")


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
    #: Lead metadata columns: ``meta={'article': True, 'unit': True}``
    #: — the financial-book form where the article number («1.1») and
    #: the unit («₸») are COLUMNS of their own before the values, not
    #: text baked into the row label. Article numbers come from each
    #: row's ``excel_props={'article': '1.1'}``.
    #: Lead metadata columns. ``True`` = the column renders with no
    #: header (historical); a STRING renders it and heads it with that
    #: text — the flag is its own caption, so the wording stays in the
    #: model and the engine stays language-neutral.
    meta_article_col: Optional[bool | str] = None
    meta_unit_col: Optional[bool | str] = None
    #: The LABEL column's header. Unlike the two above this is not a
    #: flag — the label column always exists — so only a string means
    #: anything: it heads the column that carries the row names.
    meta_label_col: Optional[str] = None
    #: The CONSTANTS column's header («Значение»). The column itself is
    #: content-driven — it exists whenever a period-headed sheet also
    #: carries non-periodic lines — so like the label column this is
    #: caption-only: a string heads it, absence keeps it headerless.
    meta_constants_col: Optional[str] = None
    #: EXTRA lead columns, one per projected field:
    #: ``meta={'fields': [('description', 'Комментарий')]}`` prints
    #: each row's ``description`` in a column of its own, headed
    #: «Комментарий». Same flag-IS-caption rule as the columns above —
    #: ``True`` instead of a string gives the column no header.
    #: Declaration order is column order; they sit after the unit
    #: column and before the values.
    #:
    #: A list of PAIRS, not free-standing keys, because a view is a
    #: real MV tree and a field named ``description`` would collide
    #: with ``MultiVariable.description`` the moment it became an
    #: attribute of the meta node.
    meta_fields: tuple = ()
    #: Nested-section layout: ``nesting={'gap_rows': 0, 'indent': 2}``
    #: — blank rows between sections (None = today's single gap) and
    #: label indent per nesting level (None = no indent). The dense
    #: financial-book form: children directly under their header,
    #: hierarchy read from the indent, not from whitespace rows.
    nesting_gap_rows: Optional[int] = None
    nesting_indent: Optional[int] = None
    #: Default number formats by value shape: ``{'number': '#,##0'}``
    #: applies to numeric value cells that declare no explicit
    #: ``number_format`` of their own — the financial-book convention
    #: (thousands separators, red parenthesised negatives) declared
    #: ONCE and cascaded, never repeated per line.
    number_formats: Optional[Dict[str, str]] = None
    format_by_type: Any = None
    time_header: Optional[bool] = None
    time_label_format: Optional[str] = None
    #: Subtotal columns interleaved with the native periods — the book
    #: form «янв фев мар (1 кв) … (год)». ``timeline={'totals':
    #: ['quarter', 'year']}``: after each quarter's last month a
    #: quarter-total column, after each year a year-total column, each
    #: carrying a LIVE formula per line (SUM/AVERAGE/last from the
    #: line's regrain rule; a formula line gets its own formula over
    #: the operands' total cells — exactly what a hand-built book
    #: does).
    time_totals: tuple = ()
    #: The word the total columns wear on the month row — «Итого Q1»,
    #: «Итого 2026». The model declares its own language
    #: (``timeline={'totals_word': 'Итого'}``); the engine's default
    #: stays English like the month labels beside it.
    time_totals_word: Optional[str] = None
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

    The first positional is a display name — a NAMED view
    (``mo.ExcelView('Квартальный', grain='quarter')``) declared on a model
    is an inert, selectable recipe: it changes nothing until a surface
    applies it. Only the ``default_excel_view`` attribute slot enters the
    resolve cascade.
    """

    _TRACK_MODES = ("blend", "rows", "compare")

    def __init__(
        self,
        display_name: Optional[str] = None,
        *,
        orient: Optional[str] = None,
        font: Any = None,
        bands: Any = None,
        cell_types: Any = None,
        timeline: Any = None,
        tracks: Optional[str] = None,
        grain: Optional[str] = None,
        formats: Any = None,
        nesting: Any = None,
        meta: Any = None,
    ) -> None:
        super().__init__(display_name or "ExcelView")
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
            ("formats", formats),
            ("nesting", nesting),
            ("meta", meta),
        ):
            if value is not None:
                setattr(self, name, _to_node(name, value))

    def _grouped(self) -> Dict[str, Any]:
        return {f: _node_value(self._components.get(f)) for f in _GROUPS}

    def __setattr__(self, name: str, value: Any) -> None:
        # Any mutation (field set, sub-view adoption) invalidates the
        # cached snapshot below — the cache must never outlive an edit.
        self.__dict__.pop("_snapshot_cache", None)
        super().__setattr__(name, value)

    def snapshot(self) -> ResolvedView:
        """Compile this grouped view into the flat :class:`ResolvedView`.

        CACHED on the instance: renderers resolve the view cascade for
        every Variable (three times each, in fact), and every resolve
        re-walked this view's whole sub-tree — 16 527 snapshot builds
        per run on a 36-month model, ~60% of the warm run. Views don't
        mutate during a run (runtime is built after exec), so the first
        build serves them all; ``__setattr__`` drops the cache on any
        real edit.
        """
        cached = self.__dict__.get("_snapshot_cache")
        if cached is not None:
            return cached
        g = self._grouped()
        timeline = g.get("timeline") or {}
        cell_types = g.get("cell_types") or {}
        nesting = g.get("nesting") or {}
        meta = g.get("meta") or {}
        resolved = ResolvedView(
            orient=g.get("orient"),
            tracks=g.get("tracks"),
            grain=g.get("grain"),
            base_font=g.get("font"),
            band_styles=g.get("bands"),
            number_formats=g.get("formats"),
            nesting_gap_rows=nesting.get("gap_rows"),
            nesting_indent=nesting.get("indent"),
            meta_article_col=meta.get("article"),
            meta_unit_col=meta.get("unit"),
            meta_label_col=meta.get("label"),
            meta_constants_col=meta.get("constants"),
            meta_fields=tuple(
                (str(f[0]), f[1]) for f in (meta.get("fields") or ())
                if isinstance(f, (tuple, list)) and len(f) == 2 and f[1]
            ),
            format_by_type=cell_types.get("on"),
            type_colors=cell_types.get("colors"),
            time_header=timeline.get("header"),
            time_label_format=timeline.get("label_format"),
            time_totals=tuple(
                k for k in (timeline.get("totals") or ())
                if k in ("quarter", "year")
            ),
            time_totals_word=timeline.get("totals_word"),
        )
        self.__dict__["_snapshot_cache"] = resolved
        return resolved

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
