# SPDX-License-Identifier: Apache-2.0
"""Unit — measurement units with algebra for financial modeling.

A Unit represents a measurement unit like ``$``, ``hr``, ``$/hr``, or
``$/(hr * Gb)``. Units support algebra (multiply, divide, raise to an
integer power, cancel) and auto-propagate through Variable arithmetic.

Usage::

    # Parse from string
    usd = Unit('$')
    hourly_rate = Unit('$/hr')
    complex_rate = Unit('$/(hr * Gb)')

    # Algebra
    Unit('$') / Unit('hr')         # → '$/hr'
    Unit('$/hr') * Unit('hr')      # → '$'    (cancels)
    Unit('hr') ** 2                # → 'hr^2'

    # On Variables — propagates through arithmetic automatically
    revenue = Variable(1_000_000, unit='$')
    hours = Variable(160, unit='hr')
    rate = revenue / hours          # rate._unit → '$/hr'

A string names a dimension: any word the author writes (``'USD m'``,
``'GWh'``, ``'EUR/kWh'``) is one, and addition holds strictly between
dimensions. A word that labels a pure number (a discount factor, a
share, a flag) is declared once as a labelled dimensionless unit and
passed as the object — it prints its word and is inert in the algebra::

    FACTOR = Unit.dimensionless('factor')
    df = Variable(0.91, unit=FACTOR)         # prints 'factor'
    pv = cash * df                           # cash's unit: the label is inert

Four words are labelled pure numbers by convention, as strings:
``'%'``, ``'x'``, ``'bps'`` and ``'pp'``.

A unit is a plain value: assigned to a model it is an attribute, never
a row or a section.
"""

from typing import Dict, Optional


class Unit:
    """Measurement unit with symbolic algebra.

    Internally stores a dict of base symbols to integer exponents::

        '$'           → {'$': 1}
        'hr'          → {'hr': 1}
        '$/hr'        → {'$': 1, 'hr': -1}
        '$/(hr * Gb)' → {'$': 1, 'hr': -1, 'Gb': -1}
        '%'           → {} (dimensionless, labelled '%')

    Equality is the dimension alone, so every dimensionless unit equals
    every other (``'%' + 'pp'`` adds); the label is what prints.
    """

    #: The words a string names as a labelled pure number, not a
    #: dimension — by convention, the same for every author.
    _DIMENSIONLESS_SYMBOLS = frozenset({'%', 'x', 'bps', 'pp'})

    def __init__(self, symbol: Optional[str] = None, *,
                 _components: Optional[Dict[str, int]] = None):
        self._unit_components: Dict[str, int] = {}
        self._display_symbol: Optional[str] = None

        if _components is not None:
            self._unit_components = {k: v for k, v in _components.items() if v != 0}
        elif symbol is not None:
            if not isinstance(symbol, str):
                raise TypeError(
                    f"Unit takes a unit string ('USD m', '$/hr'); got "
                    f"{type(symbol).__name__}. A pure number with a label is "
                    f"Unit.dimensionless('factor')."
                )
            if symbol in self._DIMENSIONLESS_SYMBOLS:
                self._display_symbol = symbol
            else:
                self._unit_components = self._parse(symbol)

    @staticmethod
    def _parse(text: str) -> Dict[str, int]:
        """
        Parse a unit string into components dict.
        
        Supports:
            '$'             → {'$': 1}
            '$/hr'          → {'$': 1, 'hr': -1}
            'hr * Gb'       → {'hr': 1, 'Gb': 1}
            '$/(hr * Gb)'   → {'$': 1, 'hr': -1, 'Gb': -1}
            '$ * hr / Gb'   → {'$': 1, 'hr': 1, 'Gb': -1}
        """
        text = text.strip()
        if not text:
            return {}

        components: Dict[str, int] = {}

        numerator_part = text
        denominator_part = None

        if '/' in text:
            slash_idx = text.index('/')
            numerator_part = text[:slash_idx].strip()
            denominator_part = text[slash_idx + 1:].strip()

        if numerator_part:
            for sym in Unit._split_product(numerator_part):
                components[sym] = components.get(sym, 0) + 1

        if denominator_part:
            denominator_part = denominator_part.strip('()')
            for sym in Unit._split_product(denominator_part):
                components[sym] = components.get(sym, 0) - 1

        return {k: v for k, v in components.items() if v != 0}

    @staticmethod
    def _split_product(text: str) -> list:
        """Split 'hr * Gb' into ['hr', 'Gb']. Single symbol returns [symbol].
        Treats '1' as empty (dimensionless numerator in '1/hr')."""
        text = text.strip().strip('()')
        if '*' in text:
            return [s.strip() for s in text.split('*') if s.strip() and s.strip() != '1']
        if text == '1' or not text:
            return []
        return [text]

    @classmethod
    def parse(cls, text: str) -> 'Unit':
        """Parse a unit string and return a Unit instance."""
        if isinstance(text, Unit):
            return text
        return cls(text)

    @classmethod
    def dimensionless(cls, label: Optional[str] = None) -> 'Unit':
        """A dimensionless unit: a pure number, inert in the algebra.

        With a ``label`` it prints that word — in a workbook's unit
        column, beside a row's name — and multiplies, divides and adds as
        the number it is, the way ``'%'`` does::

            FACTOR = Unit.dimensionless('factor')
            pv = cash * discount       # discount in FACTOR: pv keeps cash's unit

        A string unit is always a dimension (``Unit('factor')`` is one), so
        a word that labels a pure number is declared with this, once, and
        the object is passed as the unit. Dimensionless units add to one
        another, the left label printing (a factor plus a share is a
        factor); a product of two carries no label, as with ``'%'``.
        """
        unit = cls(_components={})
        if label is None:
            return unit
        if not isinstance(label, str):
            raise TypeError(
                f"a unit's label is a word (str); got {type(label).__name__}."
            )
        word = label.strip()
        if not word:
            raise ValueError("a unit's label is a word; got an empty one.")
        if any(ch in word for ch in '/*^'):
            raise ValueError(
                f"a label is one word, not an expression: {label!r}. A unit "
                f"built of others is a dimension, not a label."
            )
        unit._display_symbol = word
        return unit

    def __mul__(self, other):
        if isinstance(other, Unit):
            merged = dict(self._unit_components)
            for k, v in other._unit_components.items():
                merged[k] = merged.get(k, 0) + v
            return Unit(_components=merged)
        return NotImplemented

    def __truediv__(self, other):
        if isinstance(other, Unit):
            merged = dict(self._unit_components)
            for k, v in other._unit_components.items():
                merged[k] = merged.get(k, 0) - v
            return Unit(_components=merged)
        return NotImplemented

    def __pow__(self, exponent):
        """Raise each component's exponent. Only integer powers are defined —
        fractional powers have no clean string rendering in Excel."""
        if not isinstance(exponent, int):
            return NotImplemented
        return Unit(_components={k: v * exponent for k, v in self._unit_components.items()})

    def __eq__(self, other):
        if isinstance(other, Unit):
            return self._unit_components == other._unit_components
        return NotImplemented

    def __hash__(self):
        return hash(tuple(sorted(self._unit_components.items())))

    def compatible_with(self, other: 'Unit') -> bool:
        """Check if two units are compatible for addition/subtraction."""
        if other is None:
            return True
        if not isinstance(other, Unit):
            return True
        return self._unit_components == other._unit_components

    @property
    def is_dimensionless(self) -> bool:
        return len(self._unit_components) == 0

    def __str__(self):
        if not self._unit_components:
            return self._display_symbol or ''

        pos = {k: v for k, v in self._unit_components.items() if v > 0}
        neg = {k: -v for k, v in self._unit_components.items() if v < 0}

        def _format_part(parts: dict) -> str:
            result = []
            for sym, exp in parts.items():
                if exp == 1:
                    result.append(sym)
                else:
                    result.append(f'{sym}^{exp}')
            return ' * '.join(result)

        num_str = _format_part(pos)
        den_str = _format_part(neg)

        if num_str and den_str:
            if len(neg) > 1:
                return f'{num_str}/({den_str})'
            return f'{num_str}/{den_str}'
        elif num_str:
            return num_str
        elif den_str:
            if len(neg) > 1:
                return f'1/({den_str})'
            return f'1/{den_str}'
        return ''

    def __repr__(self):
        # The constructor that rebuilds it: the string for a dimension or a
        # conventional word, the classmethod for a declared label.
        if self._unit_components or self._display_symbol in self._DIMENSIONLESS_SYMBOLS:
            return f"Unit({str(self)!r})"
        if self._display_symbol:
            return f"Unit.dimensionless({self._display_symbol!r})"
        return "Unit.dimensionless()"

    def __bool__(self):
        """A Unit is truthy if it has components or a display symbol (e.g. '%')."""
        return bool(self._unit_components) or bool(self._display_symbol)


def unit_key(unit: Unit) -> tuple:
    """A unit as one value for "the same unit": its dimension and, for a
    pure number, its label. Two lines in ``'%'`` and in ``'x'`` add (both
    are pure numbers), but an aggregate over them names neither."""
    components = tuple(sorted(unit._unit_components.items()))
    return (components, None if components else unit._display_symbol)
