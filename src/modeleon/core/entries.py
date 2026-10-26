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

from typing import Any, Dict, List, Optional, Set, Tuple

#: How many named formulas deep the walk reads.
MAX_DEPTH = 40

#: Readings one shared ``readings`` dict keeps per track before it starts
#: afresh - room for every line of a large model, read either way.
READINGS_LIMIT = 1 << 16

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


def entry_lines(expr: Any, track: str, *, own: bool = False,
                readings: Optional[dict] = None) -> List[Entry]:
    """The rows of entries ``expr`` reads on ``track``, in reading order,
    each once. ``own``: ``expr`` was written for this track alone
    (``actual=s.raw * 1000``), so the lines without tracks it reads - loaded
    data - are its entries too. ``readings``: see :func:`readings_under`."""
    found = _found(expr, track, False, own, readings)
    seen: Dict[Tuple[int, Optional[str]], Entry] = {}
    for line, role, _offset in found:
        seen.setdefault((id(line), role), (line, role))
    return list(seen.values())


def shown_entered_mask(expr: Any, track: str, periods: int, *,
                       own: bool = False,
                       readings: Optional[dict] = None) -> Optional[List[bool]]:
    """Per period, was this formula's month entered? A record of which
    cells of a formula line were entered, never used for its value. The
    entries of the same month decide; a formula that reads only earlier
    months (salaries paid a month late: ``lag(accrued)``) is asked of
    them, at the period each reads, and a lag's first periods - its fill,
    a known number - count as entered. None when it reads no entries.
    ``readings``: see :func:`readings_under`."""
    found = _found(expr, track, True, own, readings)
    if not found:
        return None
    from .tracks import TrackValues

    same_month = [entry for entry in found if entry[2] == 0]
    reads: List[Tuple[list, int]] = []
    read: Set[Tuple[int, int]] = set()     # a row reached twice is read once
    for line, role, offset in (same_month or found):
        value = getattr(line, "_value", None)
        if isinstance(value, TrackValues):
            value = value[role] if role is not None and role in value else None
        if isinstance(value, list) and (id(value), offset) not in read:
            read.add((id(value), offset))
            reads.append((value, offset))
    if not reads:
        return None

    # A period is entered once any read says so: each read is asked only
    # about the periods none before it has answered.
    mask = [False] * periods
    open_periods: Any = range(periods)
    for series, offset in reads:
        length = len(series)
        still: List[int] = []
        for t in open_periods:
            i = t - offset
            if i < 0 or i >= length or series[i] is not None:
                mask[t] = True
            else:
                still.append(t)
        if not still:
            break
        open_periods = still
    return mask


def readings_under(node: Any) -> Optional[dict]:
    """The readings shared by every line under ``node``'s topmost
    container, kept on that container (None when it keeps no attributes).

    Every line of a model is asked what it reads as it joins, and a total
    reads its subtotals, which read the lines below: walked afresh for
    each line, the same subtotals are read over and over. A reading is
    kept per line, track and way of reading, and reused only while the
    line - and every line it reads through - still holds the very value,
    formula and track bindings it was read from (they are replaced, never
    changed in place, when a line is rebound, re-materialized or given a
    new definition). So whatever the model does between two questions,
    the answer is the one a fresh walk gives; the dict only lives as long
    as the model does."""
    top = node
    seen: Set[int] = set()
    while id(top) not in seen:
        seen.add(id(top))
        above = getattr(top, '_owner', None) or getattr(top, '_parent', None)
        if above is None:
            break
        top = above
    slots = getattr(top, '__dict__', None)
    if not isinstance(slots, dict):
        return None
    return slots.setdefault('_entry_readings', {})


def _found(expr: Any, track: str, shifts: bool, own: bool,
           readings: Optional[dict]) -> List[_Shifted]:
    """What ``expr`` reads on ``track``: through the shared ``readings``
    when given; afresh when there are none, or when the walk would stop at
    ``MAX_DEPTH`` (what it finds there depends on the way it came)."""
    if readings is not None:
        try:
            return _Walk(track, shifts, readings).formula(expr, 0, own)
        except _TooDeep:
            pass
    return _Walk(track, shifts).formula(expr, 0, own)


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


class _TooDeep(Exception):
    """A walk that keeps readings would stop at ``MAX_DEPTH`` here - a chain
    of formulas that long, or lines that read each other."""


class _Reading:
    """What one line reads on one track, with what it was read from: the
    line's value, formula and track bindings, the track's operand (its
    owner and formula decide whether the track is a reference), each line
    its formula names - with the formula read in place when it was one
    (an unnamed step) - and the readings of the lines it read through."""

    __slots__ = ('line', 'value', 'expr', 'kwargs', 'operand', 'operand_state',
                 'refs', 'parts', 'found', 'height', 'too_deep')

    def __init__(self, line: Any, value: Any, expr: Any, kwargs: Any) -> None:
        self.line = line
        self.value = value
        self.expr = expr
        self.kwargs = kwargs
        self.operand: Any = None
        self.operand_state: Optional[Tuple[bool, Any]] = None
        self.refs: List[Tuple[Any, Any]] = []
        self.parts: List['_Reading'] = []
        self.found: List[_Shifted] = []
        #: Named formulas read through, down the longest way (0: none).
        self.height = 0
        #: Met too deep to read through: ``height`` is then how far down
        #: the line reads at least, and the reading answers nothing.
        self.too_deep = False

    @classmethod
    def deeper_than(cls, line: Any, height: int) -> '_Reading':
        """A line a walk gave up on: it reads at least ``height`` named
        formulas down. A later walk meeting it as deep gives up at once;
        one from higher up reads it again."""
        reading = cls(line, getattr(line, "_value", None), getattr(line, "_expr", None),
                      getattr(line, "_role_kwargs", None))
        reading.height = height
        reading.too_deep = True
        return reading


class _Walk:
    def __init__(self, track: str, shifts: bool,
                 readings: Optional[dict] = None) -> None:
        self.track = track
        self.shifts = shifts              # read through a lag, at its offset
        if readings is None:
            self.memo: Dict[Tuple[int, bool], _Reading] = {}
            self.kept: Optional[Set[int]] = None
        else:
            self.memo = readings.setdefault((track, shifts), {})
            if len(self.memo) > READINGS_LIMIT:
                self.memo.clear()
            self.kept = set()             # readings found still true this walk
        self.open: Set[Tuple[int, bool]] = set()

    def formula(self, expr: Any, depth: int, own: bool = False,
                reading: Optional[_Reading] = None) -> List[_Shifted]:
        from .expr import (BinOp, Compare, FuncCall, Literal, MethodCall, Paren,
                           Restrict, UnaryOp, VarRef)

        out: List[_Shifted] = []
        stack = [(expr, 0)]
        seen: set = set()
        while stack:
            node, offset = stack.pop()
            # Each kind of node is its own class (none is another's
            # subclass), so the order of the tests is free: the common
            # ones - a reference, a literal, arithmetic - come first.
            if isinstance(node, VarRef):
                var = node.var
                if (id(var), offset) in seen:
                    continue
                seen.add((id(var), offset))
                inner = getattr(var, "_expr", None)
                inline = getattr(var, "_owner", None) is None and inner is not None
                if reading is not None:
                    reading.refs.append((var, inner if inline else None))
                if inline:
                    stack.append((inner, offset))   # unnamed - part of the formula
                else:
                    part = self.line(var, depth, own)
                    if reading is not None:
                        reading.parts.append(part)
                    out.extend((line, role, offset + o)
                               for line, role, o in part.found)
            elif isinstance(node, Literal):
                pass                              # a constant: no entry
            elif isinstance(node, (BinOp, Compare)):
                stack.extend(((node.right, offset), (node.left, offset)))
            elif isinstance(node, Paren):
                stack.append((node.inner, offset))
            elif isinstance(node, FuncCall):
                if node.func.upper() in SAME_MONTH_FUNCTIONS:
                    stack.extend((a, offset) for a in reversed(node.args))
            elif isinstance(node, UnaryOp):
                stack.append((node.operand, offset))
            elif isinstance(node, MethodCall):
                if node.method == 'copy':
                    stack.append((node.base, offset))
                elif (node.method == 'shift' and self.shifts and node.args
                      and isinstance(node.args[0], int)):
                    stack.append((node.base, offset + node.args[0]))
            elif isinstance(node, Restrict):
                # ``line.at(track=...)`` reads that one track: this track's
                # entries when it names this track, none when another.
                if node.label == self.track:
                    stack.append((node.base, offset))
        return out

    def line(self, var: Any, depth: int, own: bool) -> _Reading:
        key = (id(var), own)
        reading = self.memo.get(key)
        if self.kept is None:
            if reading is None:
                reading = self.memo[key] = self._line(var, depth, own)
            return reading
        if reading is not None and reading.too_deep:
            if (reading.line is var and depth + reading.height > MAX_DEPTH
                    and self._holds(reading)):
                raise _TooDeep
            reading = None
        if reading is None or reading.line is not var or not self._holds(reading):
            if key in self.open:
                raise _TooDeep            # a line that reads itself
            self.open.add(key)
            try:
                reading = self._line(var, depth, own)
            except _TooDeep:
                # Something below stops the walk MAX_DEPTH down: the line
                # reads at least that far from here. Remember it, so the
                # next walk meeting it as deep does not read down again.
                self.memo[key] = _Reading.deeper_than(var, MAX_DEPTH - depth + 1)
                raise
            finally:
                self.open.discard(key)
            self.memo[key] = reading
            self.kept.add(id(reading))
        if depth + reading.height > MAX_DEPTH:
            raise _TooDeep
        return reading

    def _holds(self, reading: _Reading) -> bool:
        """Is ``reading`` still what its line reads? Its line, the lines
        its formula names and the lines below hold the same state."""
        kept = self.kept
        assert kept is not None
        if id(reading) in kept:
            return True
        var = reading.line
        if (getattr(var, "_value", None) is not reading.value
                or getattr(var, "_expr", None) is not reading.expr
                or getattr(var, "_role_kwargs", None) is not reading.kwargs):
            return False
        if reading.operand_state is not None:
            operand = reading.operand
            if ((getattr(operand, "_owner", None) is None) is not reading.operand_state[0]
                    or getattr(operand, "_expr", None) is not reading.operand_state[1]):
                return False
        for ref, inlined in reading.refs:
            inner = getattr(ref, "_expr", None)
            inline = getattr(ref, "_owner", None) is None and inner is not None
            if (inner if inline else None) is not inlined:
                return False
        for part in reading.parts:
            if not self._holds(part):
                return False
        kept.add(id(reading))
        return True

    def _line(self, var: Any, depth: int, own: bool) -> _Reading:
        from .expr import ListExpr
        from .tracks import TrackValues

        value = getattr(var, "_value", None)
        expr = getattr(var, "_expr", None)
        kwargs = getattr(var, "_role_kwargs", None)
        reading = _Reading(var, value, expr, kwargs)
        if not isinstance(value, TrackValues):
            # A line without tracks is the same on every track: no actual
            # of its own - unless the expression was written for this
            # track, where it is what the actual is made of.
            if not own or not isinstance(value, list):
                return reading
            if expr is None:
                reading.found = [(var, None, 0)]
                return reading
            through, through_own = expr, True
        else:
            role = self.track
            if role not in value.roles or not isinstance(value[role], list):
                return reading            # a constant is no month's entry
            operand = (kwargs or {}).get(role)
            if isinstance(operand, list):
                # Entered as a list - of numbers, or a row of references
                # (``track_expr`` is None or a list expression): cell by
                # cell, blank where nothing was. Asking ``track_expr`` would
                # only look the cells over.
                reading.found = [(var, role, 0)]
                return reading
            if operand is not None:
                # A line as the operand is a reference unless it is a typed
                # value (no owner, no formula or a typed date).
                reading.operand = operand
                reading.operand_state = (getattr(operand, "_owner", None) is None,
                                         getattr(operand, "_expr", None))
            authored = var.track_expr(role)
            if isinstance(authored, ListExpr):
                reading.found = [(var, role, 0)]
                return reading            # cell by cell: blank where nothing was
            through = authored if authored is not None else expr
            if through is None or (authored is None and role in (kwargs or {})):
                reading.found = [(var, role, 0)]
                return reading            # entered: blank where nothing was
            through_own = authored is not None
        if depth >= MAX_DEPTH:
            if self.kept is not None:
                raise _TooDeep
            return reading
        reading.found = self.formula(through, depth + 1, through_own,
                                     reading if self.kept is not None else None)
        reading.height = 1 + max((part.height for part in reading.parts), default=0)
        return reading
