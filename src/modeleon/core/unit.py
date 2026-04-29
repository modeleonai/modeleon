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

``Unit`` inherits from :class:`MultiVariableClass` so unit instances
can be placed in the model tree as nodes — this keeps the seam open
for future unit-conversion support (FX rates, hr ↔ day, $ ↔ $K)
without a breaking refactor. Today ``Unit`` only carries its symbol
algebra; the MV facilities are unused.
"""

from typing import Dict, Optional

from .multi_variable import MultiVariableClass


class Unit(MultiVariableClass):
    """Measurement unit with symbolic algebra.

    Internally stores a dict of base symbols to integer exponents::

        '$'           → {'$': 1}
        'hr'          → {'hr': 1}
        '$/hr'        → {'$': 1, 'hr': -1}
        '$/(hr * Gb)' → {'$': 1, 'hr': -1, 'Gb': -1}
        '%'           → {} (dimensionless)

    Inherits from :class:`MultiVariableClass` for forward compatibility
    with unit-conversion (see module docstring); none of the MV
    machinery is exercised here today.
    """

    _DIMENSIONLESS_SYMBOLS = {'%', 'x', 'bps', 'pp'}

    def __init__(self, symbol: Optional[str] = None, *, _components: Optional[Dict[str, int]] = None, **kwargs):
        self._unit_components: Dict[str, int] = {}
        self._display_symbol: Optional[str] = None

        if _components is not None:
            self._unit_components = {k: v for k, v in _components.items() if v != 0}
        elif symbol is not None:
            if symbol in self._DIMENSIONLESS_SYMBOLS:
                self._unit_components = {}
                self._display_symbol = symbol
            else:
                self._unit_components = self._parse(symbol)

        super().__init__(**kwargs)

    def compute(self, **kwargs):
        from ..core.variable import Variable
        self.display = Variable(str(self), value_type='string')

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
        if text in cls._DIMENSIONLESS_SYMBOLS:
            return cls(text)
        return cls(text)

    @classmethod
    def dimensionless(cls) -> 'Unit':
        """Return a dimensionless unit (empty components)."""
        return cls(_components={})

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
        return f"Unit('{self}')"

    def __bool__(self):
        """A Unit is truthy if it has components or a display symbol (e.g. '%')."""
        return bool(self._unit_components) or bool(self._display_symbol)
