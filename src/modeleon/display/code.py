# SPDX-License-Identifier: Apache-2.0
"""Runtime → DSL source renderer.

Emits Modeleon Python source from a live :class:`~modeleon.Variable`,
:class:`~modeleon.MultiVariable`, or :class:`~modeleon.Model`. Sibling
to :func:`~modeleon.display.html.variable_html` / :func:`model_html`
and the JSON renderer — same idea, different target representation.

The output is **self-relative**: calling ``.code`` on any object
treats that object as the root of the snippet. ``model.code`` emits
``model = mo.Model(...)`` followed by every descendant qualified
under ``model.``; ``mv.code`` does the same starting at ``mv``;
``variable.code`` emits a single ``name = mo.Variable(...)`` line.

Formula vs value (Excel-shaped). For each Variable being emitted,
the renderer chooses between the formula form and the value form by
checking whether every reference in the formula points to a Variable
that's *also* in the snippet's scope. References that escape the
scope can't be resolved when the snippet runs in isolation, so the
renderer falls back to the computed value — same pattern as the
Formulas/Values toggle in the HTML repr.

Lossy w.r.t. user-written ``.py`` — arbitrary code between Variable
declarations (imports, helpers, comments) lives in source bytes and
is invisible to runtime. ``.code`` reproduces the DSL subset; not
every line you might have written.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterable

if TYPE_CHECKING:
    from ..core.expr import Expr
    from ..core.multi_variable import MultiVariableBase
    from ..core.variable import Variable


# ─── Public surface ──────────────────────────────────────────────


def render_code(obj: "Variable | MultiVariableBase") -> str:
    """Dispatch to the right renderer based on the node's kind.

    Single entry point used by the ``Component.code`` property — keeps
    Variable and MultiVariable on the same calling convention even
    though the rendered shape is different (one line vs full subtree).
    """
    from ..core.multi_variable import MultiVariableBase
    from ..core.variable import Variable

    if isinstance(obj, Variable):
        return variable_code(obj)
    if isinstance(obj, MultiVariableBase):
        return mv_code(obj)
    raise TypeError(
        f"code rendering not supported for {type(obj).__name__}"
    )


def variable_code(var: "Variable") -> str:
    """Render ``var`` as a single ``name = mo.Variable(...)`` line.

    The LHS is ``var.python_name`` when the Variable has been adopted
    (assigned via ``parent.x = mo.Variable(...)``); otherwise falls back
    to the path leaf. The RHS is the formula when the Variable's
    references are all in scope (here: only literal-RHS Variables, since
    a single Variable has no siblings), otherwise the computed value.

    No trailing newline.
    """
    return _render_variable(
        var,
        lhs=_default_lhs(var),
        root_lhs=None,
        root_segments=(),
        in_scope=set(),
    )


def mv_code(mv: "MultiVariableBase") -> str:
    """Render ``mv`` plus every descendant as a runnable snippet.

    The header line declares ``mv`` itself (``mo.Model(...)`` for a
    :class:`~modeleon.Model`, ``mo.MultiVariable(...)`` otherwise).
    Descendants follow with LHSes qualified under ``mv``'s python_name —
    e.g. ``mv.child = mo.Variable(...)``, ``mv.sub.deep = ...``.

    Variable formulas keep their original form when every reference
    points at another Variable inside this subtree; otherwise they
    fall back to the value form so the snippet runs standalone. Refs
    inside formulas are emitted as full attribute paths
    (``co.org.dept.headcount``) so the script resolves names at
    exec-time regardless of nesting depth.

    Trailing newline included.
    """
    root_lhs = _default_lhs(mv)
    root_segments = mv.path.segments
    in_scope = _collect_scope_vars(mv)
    lines: list[str] = [_render_mv_header(mv, lhs=root_lhs)]
    _emit_descendants(
        mv,
        prefix=root_lhs,
        root_lhs=root_lhs,
        root_segments=root_segments,
        in_scope=in_scope,
        out=lines,
    )
    return "\n".join(lines) + "\n"


# ─── Render: single Variable ─────────────────────────────────────


def _render_variable(
    var: "Variable",
    lhs: str,
    root_lhs: str | None,
    root_segments: tuple[str, ...],
    in_scope: set["Variable"],
) -> str:
    rhs = _render_rhs(var, root_lhs, root_segments, in_scope)
    kwargs = _render_kwargs(var)
    return f"{lhs} = mo.Variable({rhs}{kwargs})"


def _render_rhs(
    var: "Variable",
    root_lhs: str | None,
    root_segments: tuple[str, ...],
    in_scope: set["Variable"],
) -> str:
    """Formula form (with VarRefs qualified by their full path under
    ``root_lhs``) when every ref is renderable — either in scope, or
    inlinable as a literal value. Otherwise fall back to the
    computed value."""
    if var._expr is not None and _refs_all_renderable(var._expr, in_scope):
        return _render_expr_qualified(
            var._expr, root_lhs, root_segments, in_scope
        )
    if (
        var._expr is None
        and var._raw_formula_str is not None
        and root_lhs is None
    ):
        # Raw-string formula — we can't statically rewrite VarRefs
        # inside arbitrary user text, so emit literally only when
        # we're in the top-level (no scope qualification needed
        # anyway). For subtree renders, fall through to the value
        # form.
        return var._raw_formula_str
    return _render_value(var.value)


def _refs_all_renderable(expr: "Expr", in_scope: set["Variable"]) -> bool:
    """``True`` when every ``VarRef`` either points into the
    snippet's scope (qualifies to a real attribute path) or is an
    *unnamed temporary* whose value can be inlined as a literal
    (e.g. ``revenue * mo.Variable(0.6)`` — the ``0.6`` has no
    python_name and folds into the rendered formula as the literal
    ``0.6``).

    Named variables that fall outside the snippet's scope make the
    formula non-renderable — the caller falls back to the computed
    value of the whole expression. (``m.total.code`` for a tree-leaf
    Variable referring to its siblings produces ``mo.Variable(30)``,
    not ``mo.Variable(10 + 20)``.)
    """
    for ref in expr.iter_refs():
        if ref in in_scope:
            continue
        if _is_unnamed_temporary(ref) and _is_inlinable_literal(ref.value):
            continue
        return False
    return True


def _is_unnamed_temporary(var: "Variable") -> bool:
    """A Variable with no ``python_name`` and a floating path — the
    short-lived intermediate produced by literal scalars inside
    expressions (``mo.Variable(0.6)`` used arithmetically), not a
    user-named entity."""
    return var.python_name is None and var.path.is_floating


def _is_inlinable_literal(value: object) -> bool:
    """Whether ``value`` can be rendered back as a Python literal."""
    if value is None:
        return False
    if isinstance(value, (int, float, str, bool)):
        return True
    if isinstance(value, list):
        return all(_is_inlinable_literal(v) for v in value)
    return False


def _render_expr_qualified(
    expr: "Expr",
    root_lhs: str | None,
    root_segments: tuple[str, ...],
    in_scope: set["Variable"],
) -> str:
    """Render ``expr`` as Python text. ``VarRef`` nodes are emitted
    as full attribute paths rooted at ``root_lhs`` so the formula
    resolves at exec-time regardless of nesting depth — a reference
    to ``co.org.dept.headcount`` renders exactly as
    ``co.org.dept.headcount`` inside the snippet's namespace.

    When ``root_lhs`` is None (top-level :func:`variable_code`),
    falls back to bare leaf names — matches engine ``to_string``
    behaviour for isolated Variable rendering.

    Mirrors the structure of each :class:`~modeleon.core.expr.Expr`
    node's own ``to_string`` — kept here (rather than passed in via
    the generic :class:`~modeleon.compile.walker.Walker`) because
    qualification is a code-rendering concern, not an Excel/JSON
    concern, and the walker's :class:`RenderCtx` is shaped for the
    backend renderers.
    """
    from ..core.expr import (
        BinOp, Compare, Expr, FuncCall, ListExpr, Literal, MethodCall,
        Paren, RollingAggregate, SelfRef, Subscript, UnaryOp, VarRef,
    )

    def qualify(var: "Variable") -> str:
        """Emit ``root_lhs.<relative path>`` for a Variable that's
        in scope. Strips the snippet root's path segments, joins the
        rest under ``root_lhs``."""
        if root_lhs is None:
            return var.python_name or var.path.leaf
        rel = var.path.segments[len(root_segments):]
        if not rel:
            # Defensive: shouldn't happen for a child VarRef, but if
            # it did (referencing the snippet root itself), emit the
            # bare LHS.
            return root_lhs
        return root_lhs + "." + ".".join(rel)

    def go(e: "Expr") -> str:
        if isinstance(e, VarRef):
            if e.var in in_scope:
                return qualify(e.var)
            # Out-of-scope ref — caller's ``_refs_all_renderable``
            # gate guarantees this is inlinable as a literal value
            # (typically an unnamed temporary like
            # ``mo.Variable(0.6)`` used inside a larger expression).
            return _render_value(e.var.value)
        if isinstance(e, Literal):
            return e.to_string()
        if isinstance(e, Paren):
            return f"({go(e.inner)})"
        if isinstance(e, BinOp):
            return f"{go(e.left)} {e.op} {go(e.right)}"
        if isinstance(e, UnaryOp):
            return f"{e.op}{go(e.operand)}"
        if isinstance(e, Compare):
            return f"{go(e.left)} {e.op} {go(e.right)}"
        if isinstance(e, Subscript):
            base_s = go(e.base)
            if isinstance(e.key, slice):
                start = "" if e.key.start is None else e.key.start
                stop = "" if e.key.stop is None else e.key.stop
                if e.key.step is not None:
                    return f"{base_s}[{start}:{stop}:{e.key.step}]"
                return f"{base_s}[{start}:{stop}]"
            if isinstance(e.key, str):
                return f'{base_s}["{e.key}"]'
            return f"{base_s}[{e.key}]"
        if isinstance(e, ListExpr):
            return "[" + ", ".join(go(item) for item in e.items) + "]"
        if isinstance(e, FuncCall):
            return f"{e.func}(" + ", ".join(go(a) for a in e.args) + ")"
        if isinstance(e, MethodCall):
            parts: list[str] = []
            for a in e.args:
                parts.append(go(a) if isinstance(a, Expr) else str(a))
            for k, v in e.kwargs.items():
                v_s = go(v) if isinstance(v, Expr) else str(v)
                parts.append(f"{k}={v_s}")
            return f"{go(e.base)}.{e.method}(" + ", ".join(parts) + ")"
        if isinstance(e, SelfRef):
            parts = [go(e.start), f'"{e.template}"']
            for name, ve in e.variables.items():
                parts.append(f"{name}={go(ve)}")
            if e.periods is not None:
                pstr = go(e.periods) if isinstance(e.periods, Expr) else str(e.periods)
                parts.append(f"periods={pstr}")
            return f"recurrence(" + ", ".join(parts) + ")"
        if isinstance(e, RollingAggregate):
            qual = qualify(e.source) if e.source in in_scope else (
                e.source.python_name or e.source.path.leaf
            )
            return f"{qual}.rolling_{e.func.lower()}({e.window})"
        # Unknown node — defer to its own to_string (lossy on prefix
        # qualification, but better than crashing).
        return e.to_string()

    return go(expr)


def _render_value(value: object) -> str:
    """Reproduce ``value`` as a Python literal. Lists are rendered as
    bracketed forms; scalars via :func:`repr`."""
    if isinstance(value, list):
        return "[" + ", ".join(_render_value(v) for v in value) + "]"
    return repr(value)


# ─── Render: MultiVariable header + descendants ──────────────────


def _render_mv_header(mv: "MultiVariableBase", lhs: str) -> str:
    """``lhs = mo.Model("name", kw=...)`` for a Model;
    ``lhs = mo.MultiVariable(kw=...)`` otherwise."""
    # Lazy import to avoid circular dependency: model.py imports
    # multi_variable.py, which would import this module on render.
    from ..core.model import Model

    kwargs = _render_kwargs(mv)
    if isinstance(mv, Model):
        # Model takes a required positional ``name``. ``_python_name``
        # was crystallized to that value in the constructor.
        name_arg = repr(mv._python_name or "")
        body = f"mo.Model({name_arg}{kwargs})"
    else:
        # ``_render_kwargs`` returns a leading-comma form for use as a
        # suffix; for the bare ``mo.MultiVariable(...)`` constructor
        # we strip that leading comma.
        body = f"mo.MultiVariable({kwargs.lstrip(', ')})"
    return f"{lhs} = {body}"


def _emit_descendants(
    mv: "MultiVariableBase",
    prefix: str,
    root_lhs: str,
    root_segments: tuple[str, ...],
    in_scope: set["Variable"],
    out: list[str],
) -> None:
    """Walk ``mv``'s components in declaration order, emitting one
    line per descendant (and recursing into nested MVs).

    ``prefix`` is the running qualifier for the LHS
    (``acme.org.dept``); ``root_lhs`` + ``root_segments`` describe
    the snippet's root used to render RHS formulas as full attribute
    paths.
    """
    from ..core.variable import Variable
    from ..core.multi_variable import MultiVariableBase

    for name in mv._component_order:
        child = mv._components[name]
        child_lhs = f"{prefix}.{name}"
        if isinstance(child, Variable):
            out.append(_render_variable(
                child,
                lhs=child_lhs,
                root_lhs=root_lhs,
                root_segments=root_segments,
                in_scope=in_scope,
            ))
        elif isinstance(child, MultiVariableBase):
            out.append(_render_mv_header(child, lhs=child_lhs))
            _emit_descendants(
                child,
                prefix=child_lhs,
                root_lhs=root_lhs,
                root_segments=root_segments,
                in_scope=in_scope,
                out=out,
            )


# ─── Scope analysis ──────────────────────────────────────────────


def _collect_scope_vars(mv: "MultiVariableBase") -> set["Variable"]:
    """Every Variable reachable inside ``mv``'s subtree (identity-
    based set).

    Used for the formula/value choice: a Variable's formula renders
    in formula form only when every ``VarRef`` it contains points at
    a Variable in this set — otherwise the snippet wouldn't resolve
    those references when run standalone.

    Identity-based (``set`` of objects) rather than name-based so
    nested namespaces don't collide (``co.org.dept.headcount`` vs
    ``co.hr.dept.headcount`` — same leaf name, different Variables).
    """
    from ..core.variable import Variable
    from ..core.multi_variable import MultiVariableBase

    vars_in_scope: set["Variable"] = set()

    def visit(node: "MultiVariableBase") -> None:
        for cname in node._component_order:
            child = node._components[cname]
            if isinstance(child, MultiVariableBase):
                visit(child)
            elif isinstance(child, Variable):
                vars_in_scope.add(child)

    visit(mv)
    return vars_in_scope


# ─── Helpers ─────────────────────────────────────────────────────


def _default_lhs(obj: "Variable | MultiVariableBase") -> str:
    """LHS for the snippet root: the object's python_name when
    available, falling back to its path leaf. Floating un-named
    objects render under their auto id, which is at least
    syntactically valid Python."""
    return obj.python_name or obj.path.leaf


def _render_kwargs(obj: "Variable | MultiVariableBase") -> str:
    """Emit each explicitly-set kwarg as ``, key=repr(value)``.

    Returns a leading-comma form so callers can append after a
    positional first arg. Empty string when nothing to emit.

    v0 covers ``display_name``, ``unit``, and ``keys``. ``indexed_by``
    is intentionally skipped — it usually references another
    Variable (an axis), and serializing that reference back to a
    name requires the in-scope analysis we already do for formulas;
    folding it in earns its own pass when a real consumer demands it.
    Plugin-routed kwargs and ``excel_props`` are also deferred.
    """
    from ..core.humanize import humanize_identifier

    parts: list[str] = []
    label = getattr(obj, "_display_name", None)
    if label:
        # Skip when the label matches the auto-humanized python_name —
        # ``Model('acme')`` sets ``_display_name = 'Acme'`` even though
        # the user didn't pass it; emitting that as a kwarg would be
        # noise on round-trip.
        auto = humanize_identifier(obj.python_name) if obj.python_name else None
        if label != auto:
            parts.append(f"display_name={label!r}")
    unit = getattr(obj, "_unit", None)
    if unit is not None:
        # ``_unit`` is a Unit object; its ``str()`` is the canonical
        # textual form. ``repr(str(unit))`` quotes it as a Python
        # string literal.
        parts.append(f"unit={str(unit)!r}")
    keys = getattr(obj, "_keys", None)
    if keys is not None:
        parts.append(f"keys={list(keys)!r}")
    if not parts:
        return ""
    return ", " + ", ".join(parts)
