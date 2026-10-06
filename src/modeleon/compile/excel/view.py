# SPDX-License-Identifier: Apache-2.0
"""ExcelView — a declared-once, cascading presentation view, *as a model*.

``ExcelView`` is a :class:`MultiVariable`, organised into **sub-views** (groups
of related settings, each a sub-model):

    orient                          geometry  — period axis direction
    font       {name, size}         typography — the book's own font
    bands      {section, …}         structure — section/band styling
    cell_types {on, colors{…}, styles{…}}   cell look by kind (input/reference)
    timeline   {header, label_format, opening, …}   the period-label header
    sheet      {title, gridlines, columns[…], …}    the furniture of every sheet

So a view is a real tree you navigate (``view.timeline.label_format``,
``view.cell_types.colors.input``). Renderers don't read the tree — they read a
**flat** :class:`ResolvedView` snapshot produced by :func:`resolve_excel_view`,
which walks to the nearest ``default_excel_view`` and fills unset leaves from the
global house default ``mo.default_excel_view``. Authoring/display is the grouped
model; consumption is the flat snapshot. Fields are PRESENTATION only — a view
never owns *what periods exist* (that's the model's timeline / projection;
the model computes, renderers only print).
"""

from __future__ import annotations

from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from dataclasses import fields as _dc_fields
from html import escape
from typing import Any, Dict, Iterator, Optional

from ...core.multi_variable import MultiVariable, MultiVariableBase
from ...core.variable import Variable

#: The authoring groups — the sub-views of an ExcelView.
_GROUPS = ("orient", "font", "bands", "cell_types", "timeline",
           "tracks", "grain", "formats", "nesting", "meta", "sheet")


@dataclass(frozen=True)
class ResolvedView:
    """A compiled, plain-value snapshot — the FLAT fields the renderers read.

    Groups are flattened (``timeline.header`` → ``time_header``, etc.) so the
    field-level merge over the house default inherits each leaf independently."""

    orient: Optional[str] = None
    #: How tracked lines print: 'blend' — the display-default
    #: (live) row only; 'rows' — default row + labeled per-track rows;
    #: 'compare' — period-major column groups, one column per track
    #: under each period.
    tracks: Optional[str] = None
    #: The DEFAULT track selection — which tracks a renderer shows
    #: unless told otherwise. Declared as the dict form of the group:
    #: ``tracks={'mode': 'blend', 'shown': ['actual', 'forecast',
    #: 'plan']}``. Kept on the view for custom renderers — Excel
    #: emission ignores it. None = no declared selection.
    tracks_shown: Optional[tuple] = None
    #: The view's grain — which grain this view prints/shows at.
    #: None = the model's native grain.
    grain: Optional[str] = None
    base_font: Optional[Dict[str, Any]] = None
    band_styles: Optional[Dict[str, Dict[str, Any]]] = None
    #: Lead metadata columns: ``meta={'article': True, 'unit': True}``
    #: — the financial-book form where the article number ("1.1") and
    #: the unit ("$") are COLUMNS of their own before the values, not
    #: text baked into the row label. Article numbers come from each
    #: row's ``excel_props={'article': '1.1'}``.
    #: ``True`` = the column renders with no header; a STRING renders
    #: it and heads it with that text — the flag is its own caption,
    #: so the wording stays in the model and the engine stays
    #: language-neutral.
    meta_article_col: Optional[bool | str] = None
    meta_unit_col: Optional[bool | str] = None
    #: The LABEL column's header. Unlike the two above this is not a
    #: flag — the label column always exists — so only a string means
    #: anything: it heads the column that carries the row names.
    meta_label_col: Optional[str] = None
    #: The CONSTANTS column's header ("Value"). The column itself is
    #: content-driven — it exists whenever a period-headed sheet also
    #: carries non-periodic lines — so like the label column this is
    #: caption-only: a string heads it, absence keeps it headerless.
    meta_constants_col: Optional[str] = None
    #: Line numbers by position: ``meta={'numbering': 'section'}`` numbers
    #: each first-level section's lines ``1.1``, ``1.2``, … in the order
    #: they stand — the inputs and the results; a link (a line that only
    #: repeats another) carries no number of its own, the number belongs
    #: to the line it repeats. A line's own ``article`` wins.
    meta_numbering: Optional[str] = None
    #: Which lines carry the row total (the ``('row_sum', …)`` field)
    #: without saying so each: ``meta={'row_sum_for': ['sum']}`` — a line
    #: whose re-grain rule (``'sum'``: flows, events) is listed. A line's
    #: own ``excel_props={'row_sum': ...}`` wins, ``False`` included.
    meta_row_sum_for: Optional[tuple] = None
    #: EXTRA lead columns, one per projected field:
    #: ``meta={'fields': [('description', 'Comment')]}`` prints
    #: each row's ``description`` in a column of its own, headed
    #: "Comment". Same flag-IS-caption rule as the columns above —
    #: ``True`` instead of a string gives the column no header.
    #: Declaration order is column order; they sit after the unit
    #: column and before the values.
    #:
    #: A list of PAIRS, not free-standing keys, because a view is a
    #: real MV tree and a field named ``description`` would collide
    #: with ``MultiVariable.description`` the moment it became an
    #: attribute of the meta node.
    meta_fields: Optional[tuple] = None
    #: Nested-section layout: ``nesting={'gap_rows': 0, 'indent': 2}``
    #: — blank rows between sections (None = the default single gap) and
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
    #: form "Jan Feb Mar (Q1) … (year)". ``timeline={'totals':
    #: ['quarter', 'year']}``: after each quarter's last month a
    #: quarter-total column, after each year a year-total column, each
    #: carrying a LIVE formula per line (SUM/AVERAGE/last from the
    #: line's regrain rule; a formula line gets its own formula over
    #: the operands' total cells — exactly what a hand-built book
    #: does).
    #: Unlike most fields, an unset ``totals`` is ``()`` — no totals — and
    #: a view that names no totals turns its level's totals off rather than
    #: showing a level above through: existing books depend on it (a sheet
    #: with its own view and no totals under a model whose view has them).
    time_totals: tuple = ()
    #: The word the total columns wear on the month row — "Total Q1",
    #: "Total 2026". The model declares its own language
    #: (``timeline={'totals_word': 'Gesamt'}``); the engine's default
    #: stays English like the month labels beside it.
    time_totals_word: Optional[str] = None
    #: The OPENING column — the point just before the first period,
    #: where a hand-built book keeps the day before the model starts
    #: (and so writes "start = previous end + 1" in every period, the
    #: first included). ``timeline={'opening': 'Pre-start'}`` reserves
    #: it left of the first period on every period-headed sheet, headed
    #: with that text; ``True`` leaves it headerless. The time rows fill
    #: it — the end row with the day before the start, the period
    #: number with 0 — and every other line leaves it empty.
    time_opening: Optional[bool | str] = None
    #: The anchor's own sheet shows the header rows at its top too, as
    #: references to its own rows — so every sheet opens on the same
    #: rows and freezes under them. Off: the anchor sheet shows its time
    #: rows once, where they are written.
    time_repeat_on_anchor: Optional[bool] = None
    type_colors: Optional[Dict[str, str]] = None
    #: The full look of a value cell by what it IS — ``cell_types=
    #: {'styles': {'input': {'font_color': '#0000FF', 'bg': '#FFF2CC'}}}``
    #: — in the ``excel_props`` vocabulary. ``colors`` is its font-colour
    #: shorthand; a style's own ``font_color`` wins over it, and the
    #: line's own ``excel_props`` win over both.
    type_styles: Optional[Dict[str, Dict[str, Any]]] = None
    #: A title row above the header on every sheet: the sheet's name in
    #: the label column (``'upper'`` prints it upper-cased) and its
    #: description beside it, in the column that carries the rows'
    #: descriptions when one is projected. Styled by ``bands['title']``
    #: and ``bands['title_note']``.
    sheet_title: Optional[bool | str] = None
    #: ``False`` hides the grid lines of every sheet.
    sheet_gridlines: Optional[bool] = None
    #: Every period-headed sheet keeps the constants column, holding a
    #: constant or not — so a period stands in the same column on every
    #: sheet of the book.
    sheet_same_columns: Optional[bool] = None
    #: The width and look of each column, by ROLE: ``sheet={'columns':
    #: [('label', {'width': 40}), ('unit', {'width': 9, 'font_color':
    #: '#7F7F7F'})]}``. Roles: ``article``, ``label``, ``unit``, each
    #: projected field by its name (``description``, ``row_sum``),
    #: ``constants``, ``opening``, ``period`` (every period column).
    #: ``width`` sets the column's width; the other keys dress the cells
    #: a row writes in a lead column (its article, unit and fields),
    #: under the row's own look. Pairs, like ``meta['fields']``, because
    #: a role named ``description`` must not become an attribute of the
    #: view tree.
    sheet_columns: Optional[tuple] = None
    #: A row spans the sheet: section bands and a line's fill and
    #: borders run from the first column to the last, and the row total
    #: wears the line's typography like its values do — the way a
    #: hand-built book draws a result line. Off: they cover the label
    #: and the values only.
    sheet_full_rows: Optional[bool] = None


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
    (``mo.ExcelView('Quarterly', grain='quarter')``) declared on a model
    is an inert recipe: it changes nothing until a caller
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
        sheet: Any = None,
    ) -> None:
        super().__init__(display_name or "ExcelView")
        if tracks is not None:
            # Two authoring forms: the mode string, or the dict form
            # {'mode': ..., 'shown': [...]} that also declares the
            # default track selection.
            mode = tracks.get("mode") if isinstance(tracks, dict) else tracks
            if mode is not None and mode not in self._TRACK_MODES:
                raise ValueError(
                    f"ExcelView tracks= mode must be one of "
                    f"{', '.join(self._TRACK_MODES)}; got {mode!r}."
                )
            if isinstance(tracks, dict):
                unknown = set(tracks) - {"mode", "shown"}
                if unknown:
                    raise ValueError(
                        f"ExcelView tracks= dict accepts 'mode' and "
                        f"'shown'; got {sorted(unknown)}."
                    )
                shown = tracks.get("shown")
                if shown is not None and (
                    not isinstance(shown, (list, tuple))
                    or not all(isinstance(s, str) for s in shown)
                ):
                    raise ValueError(
                        "ExcelView tracks= 'shown' must be a list of "
                        "track names."
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
            ("sheet", sheet),
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
        every Variable (three times each, in fact), and without the
        cache every resolve would re-walk this view's whole sub-tree —
        thousands of snapshot builds per render of a large model. Views
        don't mutate while a model renders, so the first build serves
        them all; ``__setattr__`` drops the cache on any real edit.
        """
        cached = self.__dict__.get("_snapshot_cache")
        if cached is not None:
            return cached
        g = self._grouped()
        timeline = g.get("timeline") or {}
        cell_types = g.get("cell_types") or {}
        nesting = g.get("nesting") or {}
        meta = g.get("meta") or {}
        sheet = g.get("sheet") or {}
        columns = sheet.get("columns")
        tracks_g = g.get("tracks")
        if isinstance(tracks_g, dict):
            tracks_mode = tracks_g.get("mode")
            shown = tracks_g.get("shown")
            tracks_shown = tuple(shown) if shown else None
        else:
            tracks_mode = tracks_g
            tracks_shown = None
        resolved = ResolvedView(
            orient=g.get("orient"),
            tracks=tracks_mode,
            tracks_shown=tracks_shown,
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
            meta_numbering=meta.get("numbering"),
            meta_row_sum_for=(
                tuple(str(k) for k in meta.get("row_sum_for") or ())
                if "row_sum_for" in meta else None
            ),
            # A group that does not mention its list leaves it UNSET
            # (None), so it inherits from the level above; only a
            # declared list — empty included — speaks for this level.
            meta_fields=(
                tuple(
                    (str(f[0]), f[1]) for f in (meta.get("fields") or ())
                    if isinstance(f, (tuple, list)) and len(f) == 2 and f[1]
                )
                if "fields" in meta else None
            ),
            format_by_type=cell_types.get("on"),
            type_colors=cell_types.get("colors"),
            type_styles=cell_types.get("styles"),
            time_header=timeline.get("header"),
            time_label_format=timeline.get("label_format"),
            time_totals=tuple(
                k for k in (timeline.get("totals") or ())
                if k in ("quarter", "year")
            ),
            time_totals_word=timeline.get("totals_word"),
            time_opening=timeline.get("opening"),
            time_repeat_on_anchor=timeline.get("repeat_on_anchor"),
            sheet_title=sheet.get("title"),
            sheet_gridlines=sheet.get("gridlines"),
            sheet_same_columns=sheet.get("same_columns"),
            sheet_columns=(
                tuple(
                    (str(c[0]), dict(c[1])) for c in columns
                    if isinstance(c, (tuple, list)) and len(c) == 2
                    and isinstance(c[1], dict)
                )
                if columns is not None else None
            ),
            sheet_full_rows=sheet.get("full_rows"),
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
    sub-view (say only ``timeline.header``) still inherits ``label_format`` etc.

    The groups that are new with the house look (the type looks, the
    formats, the sheet's columns) merge key by key, and a look style by
    style: ``formats={'date': …}`` over the floor's number formats keeps
    the others. The groups a model could declare before — the font, the
    bands, the type colours — still replace the whole group, as they
    always did: a view naming only ``colors.input`` paints nothing else,
    and a tab's ``bands`` stand alone under the model's. (Layering those
    too changed existing books: a partial ``cell_types.colors`` suddenly
    painted every cross-sheet reference the floor's green.)"""
    return ResolvedView(**{
        f: (_layered if f in _LAYERED_GROUPS else _picked)(
            getattr(snap, f), getattr(default, f)
        )
        for f in _RESOLVED_FIELDS
    })


#: The mapping groups that layer key by key (new with the house look).
_LAYERED_GROUPS = frozenset({"type_styles", "number_formats", "sheet_columns"})


def _picked(mine: Any, under: Any) -> Any:
    return under if mine is None else mine


def _layered(mine: Any, under: Any) -> Any:
    if mine is None:
        return under
    if isinstance(mine, dict) and isinstance(under, dict):
        out = dict(under)
        for key, value in mine.items():
            out[key] = _layered(value, under.get(key))
        return out
    return mine


#: The house view of the work in progress — the layer between the
#: process-wide ``mo.default_excel_view`` and the models' own views.
#: Context-local, so two threads or asyncio tasks rendering under
#: different house views at once never see each other's.
_HOUSE: ContextVar[Optional[ResolvedView]] = ContextVar("modeleon_house_view", default=None)


def changes(view: ExcelView, base: Optional[ExcelView] = None) -> ResolvedView:
    """Only what ``view`` says that ``base`` (by default the process-wide
    ``mo.default_excel_view``) does not: every field the two agree on is
    left unset. A style that began as a copy of the default and was then
    tuned — a team's house style, say — layers as the changes it made, not as
    a full copy that would reset everything beneath it."""
    mine = view.snapshot()
    floor = base if base is not None else _global_default()
    theirs = floor.snapshot() if floor is not None else ResolvedView()
    return ResolvedView(**{
        f: (getattr(mine, f) if getattr(mine, f) != getattr(theirs, f) else None)
        for f in _RESOLVED_FIELDS
    })


@contextmanager
def house_view(*views: "ExcelView | ResolvedView") -> Iterator[None]:
    """Render under a house look without touching any model::

        with mo.house_view(firm_standard, team_style):
            model.to_excel("deal.xlsx")

    Inside the block every node resolves its view over these, most
    general first: a later view's fields win over an earlier one's, a
    model's own ``default_excel_view`` wins over all of them, and
    ``mo.default_excel_view`` stays the floor under them. A view may come
    as its snapshot — :func:`changes` gives one. The layer is
    context-local (a thread or an asyncio task sees only its own), and
    it ends with the block."""
    snap: Optional[ResolvedView] = _HOUSE.get()
    for view in views:
        if isinstance(view, ExcelView):
            layer = view.snapshot()
        elif isinstance(view, ResolvedView):
            layer = view
        else:
            raise TypeError(
                f"house_view takes ExcelView objects; got {type(view).__name__}."
            )
        snap = layer if snap is None else _merge(layer, snap)
    token = _HOUSE.set(snap)
    try:
        yield
    finally:
        _HOUSE.reset(token)


def resolve_excel_view(node: Any) -> ResolvedView:
    """The effective :class:`ResolvedView` for ``node`` — resolved HIERARCHICALLY.

    Walks ``_owner`` → ``_parent`` from ``node`` up to the root, collecting EVERY
    ``default_excel_view`` pointer on the way (not just the nearest). Fields
    cascade like CSS: the *closest* level that sets a field wins, falling back up
    the tree, with the process-wide ``mo.default_excel_view`` as the top-level
    default (the floor) and the :func:`house_view` in force, if any, right
    above it. With no pointer anywhere, that floor is returned.
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
    house = _HOUSE.get()
    if house is not None:
        result = _merge(house, result)
    # Fold top-down (root-most first) so each lower level overrides its ancestors
    # and the nearest level wins; the global default stays the floor.
    for snap in reversed(chain):
        result = _merge(snap, result)
    return result
