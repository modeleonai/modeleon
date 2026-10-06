# SPDX-License-Identifier: Apache-2.0
"""Which entered rows a derived track is computed from.

A blank counts as zero in arithmetic, as in Excel, so a formula over months
nobody entered computes a number. Whether an actual was ENTERED is then a
question for the cells it is built from: a line that enters its plan and
derives its actual (``mo.Variable(a + b, plan=[...])``) has an actual in a
month when any of ``a``'s or ``b``'s actuals is entered there. The blend asks
it, and the workbook asks it the same way (``COUNTA`` of the same cells).

The walk reads arithmetic on the same month - ``+ - * /``, unary minus,
comparisons, brackets, functions that work month by month (``IF``,
``ROUND``, ...), the unnamed steps between - and the named lines with tracks it
reaches: a line of entries (no expression, or the track entered by hand) is an
entry; a formula line is read through. A line without tracks is the same on
every track, so it is no actual's entry - except inside an expression written
for one track (``actual=s.raw * 1000``), where it is what the actual is made of;
a constant never is. ``line.at(track=...)`` reads that track's entries when it
names the track asked, none otherwise. A part it cannot read (a running total,
an aggregate over time) adds no entries; a lag adds none to the blend's
question either - it reads another month - but ``shown_entered_mask`` reads
through it at its offset.
"""

from typing import Any, Dict, List, Optional, Tuple

#: How many named formulas deep the walk reads.
MAX_DEPTH = 40

#: Functions that work month by month, whose arguments' entries are the
#: result's. An aggregate (``SUM`` of a row is one number for every
#: month) or a shift in time (``lag``, ``cumsum``) is not among them.
SAME_MONTH_FUNCTIONS = frozenset({
    'IF', 'ABS', 'ROUND', 'INT', 'MOD', 'AND', 'OR', 'NOT', 'CHOOSE',
})

#: ``(line, track)`` - ``track`` is None for a line without tracks, an
#: entry only inside an expression written for one track.
Entry = Tuple[Any, Optional[str]]

#: An entry read ``offset`` periods back (``lag``): ``(line, track, offset)``.
_Shifted = Tuple[Any, Optional[str], int]


def entry_lines(expr: Any, track: str, *, own: bool = False) -> List[Entry]:
    """The rows of entries ``expr`` reads on ``track``, in reading order,
    each once. ``own``: ``expr`` was written for this track alone
    (``actual=s.raw * 1000``), so the lines without tracks it reads - loaded
    data - are its entries too."""
    found = _Walk(track, shifts=False).formula(expr, 0, own)
    seen: Dict[Tuple[int, Optional[str]], Entry] = {}
    for line, role, _offset in found:
        seen.setdefault((id(line), role), (line, role))
    return list(seen.values())


def shown_entered_mask(expr: Any, track: str, periods: int, *,
                       own: bool = False) -> Optional[List[bool]]:
    """Per period, was this formula's month entered? A record of which
    cells of a formula line were entered, never used for its value. The
    entries of the same month decide; a formula that reads only earlier
    months (salaries paid a month late: ``lag(accrued)``) is asked of
    them, at the period each reads, and a lag's first periods - its fill,
    a known number - count as entered. None when it reads no entries."""
    found = _Walk(track, shifts=True).formula(expr, 0, own)
    if not found:
        return None
    from .tracks import TrackValues

    same_month = [entry for entry in found if entry[2] == 0]
    reads = []
    for line, role, offset in (same_month or found):
        value = getattr(line, "_value", None)
        if isinstance(value, TrackValues):
            value = value[role] if role is not None and role in value.roles else None
        if isinstance(value, list):
            reads.append((value, offset))
    if not reads:
        return None

    def entered(t: int) -> bool:
        for series, offset in reads:
            i = t - offset
            if i < 0 or i >= len(series) or series[i] is not None:
                return True
        return False

    return [entered(t) for t in range(periods)]


def entered_mask(entries: List[Entry], periods: int) -> List[bool]:
    """Per period: is any of the entries filled there?"""
    from .tracks import TrackValues

    series = []
    for line, role in entries:
        value = getattr(line, "_value", None)
        if isinstance(value, TrackValues):
            value = value[role] if role is not None and role in value.roles else None
        if isinstance(value, list):
            series.append(value)
    return [any(t < len(s) and s[t] is not None for s in series)
            for t in range(periods)]


class _Walk:
    def __init__(self, track: str, shifts: bool) -> None:
        self.track = track
        self.shifts = shifts              # read through a lag, at its offset
        self.memo: Dict[Tuple[int, bool], List[_Shifted]] = {}

    def formula(self, expr: Any, depth: int, own: bool = False) -> List[_Shifted]:
        from .expr import (BinOp, Compare, FuncCall, MethodCall, Paren, Restrict,
                           UnaryOp, VarRef)

        out: List[_Shifted] = []
        stack = [(expr, 0)]
        seen: set = set()
        while stack:
            node, offset = stack.pop()
            if isinstance(node, Restrict):
                # ``line.at(track=...)`` reads that one track: this track's
                # entries when it names this track, none when another.
                if node.label == self.track:
                    stack.append((node.base, offset))
            elif isinstance(node, Paren):
                stack.append((node.inner, offset))
            elif isinstance(node, UnaryOp):
                stack.append((node.operand, offset))
            elif isinstance(node, (BinOp, Compare)):
                stack.extend(((node.right, offset), (node.left, offset)))
            elif isinstance(node, FuncCall):
                if node.func.upper() in SAME_MONTH_FUNCTIONS:
                    stack.extend((a, offset) for a in reversed(node.args))
            elif isinstance(node, MethodCall):
                if node.method == 'copy':
                    stack.append((node.base, offset))
                elif (node.method == 'shift' and self.shifts and node.args
                      and isinstance(node.args[0], int)):
                    stack.append((node.base, offset + node.args[0]))
            elif isinstance(node, VarRef):
                var = node.var
                if (id(var), offset) in seen:
                    continue
                seen.add((id(var), offset))
                if (getattr(var, "_owner", None) is None
                        and getattr(var, "_expr", None) is not None):
                    stack.append((var._expr, offset))   # unnamed - part of the formula
                else:
                    out.extend((line, role, offset + o)
                               for line, role, o in self.line(var, depth, own))
        return out

    def line(self, var: Any, depth: int, own: bool) -> List[_Shifted]:
        key = (id(var), own)
        if key not in self.memo:
            self.memo[key] = self._line(var, depth, own)
        return self.memo[key]

    def _line(self, var: Any, depth: int, own: bool) -> List[_Shifted]:
        from .expr import ListExpr
        from .tracks import TrackValues

        value = getattr(var, "_value", None)
        if not isinstance(value, TrackValues):
            # A line without tracks is the same on every track: no actual
            # of its own - unless the expression was written for this
            # track, where it is what the actual is made of.
            if not own or not isinstance(value, list):
                return []
            expr = getattr(var, "_expr", None)
            if expr is None:
                return [(var, None, 0)]
            return self.formula(expr, depth + 1, True) if depth < MAX_DEPTH else []
        role = self.track
        if role not in value.roles or not isinstance(value[role], list):
            return []                     # a constant is no month's entry
        authored = var.track_expr(role)
        if isinstance(authored, ListExpr):
            return [(var, role, 0)]       # cell by cell: blank where nothing was
        expr = authored if authored is not None else getattr(var, "_expr", None)
        if expr is None or (authored is None and role in (
                getattr(var, "_role_kwargs", None) or {})):
            return [(var, role, 0)]       # entered: blank where nothing was
        if depth >= MAX_DEPTH:
            return []
        return self.formula(expr, depth + 1, authored is not None)
