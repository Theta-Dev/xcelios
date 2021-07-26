import re
from enum import Enum
from typing import Tuple

from openpyxl.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

_POSITION_RE = re.compile(r'([^!]+)!\$([A-Z]+)\$(\d+)')
_COORD_RE = re.compile(r'([A-Z]+)(\d+)')


class InvalidPositionError(Exception):
    pass


class Direction(Enum):
    UP = (0, -1)
    RIGHT = (1, 0)
    DOWN = (0, 1)
    LEFT = (-1, 0)

    def __str__(self):
        return self.name.lower()


class Position:
    def __init__(self, *args):
        if len(args) == 2:
            self.col = Position._parse_colval(args[0])
            self.row = Position._parse_rowval(args[1])
        elif len(args) == 1:
            m = re.match(_COORD_RE, args[0])

            if not m:
                raise InvalidPositionError('Invalid coord string: %s' %
                                           args[0])

            self.col = Position._parse_colval(m[1])
            self.row = Position._parse_rowval(m[2])
        else:
            raise TypeError('Position requires 1-2 positional arguments')

    @staticmethod
    def _parse_colval(val: [int, str]) -> int:
        if not isinstance(val, int):
            try:
                val = column_index_from_string(val)
            except ValueError:
                raise InvalidPositionError('Invalid column index: %s' %
                                           str(val))

        if not 1 <= val <= 16384:
            raise InvalidPositionError('Column index out of range: %d' % val)

        return val

    @staticmethod
    def _parse_rowval(val: [int, str]) -> int:
        if not isinstance(val, int):
            try:
                val = int(val)
            except ValueError:
                raise InvalidPositionError('Invalid row index: %s' % str(val))

        if not 1 <= val <= 1048576:
            raise InvalidPositionError('Row index out of range: %d' % val)

        return val

    @classmethod
    def from_abs_string(cls, abs_str: str) -> Tuple['Position', str]:
        m = re.match(_POSITION_RE, abs_str)

        if not m:
            raise InvalidPositionError('Invalid position string: %s' % abs_str)

        return cls(m[2], m[3]), m[1]

    def is_in(self, ws: Worksheet) -> bool:
        return ws.min_row <= self.row <= ws.max_row and \
               ws.min_column <= self.col <= ws.max_column

    def get_cell(self, ws: Worksheet) -> Cell:
        return ws.cell(self.row, self.col)

    def shifted(self, direction: Direction, d: int = 1) -> 'Position':
        ncol = self.col + direction.value[0] * d
        nrow = self.row + direction.value[1] * d

        return Position(ncol, nrow)

    def __eq__(self, other):
        return self.row == other.row and self.col == other.col

    def __str__(self):
        return get_column_letter(self.col) + str(self.row)

    def __repr__(self):
        return '<Position: %s>' % str(self)


class MarkerAbs:
    def get_position(self, ws: Worksheet) -> Position:
        pass

    def get_cell(self, ws: Worksheet) -> Cell:
        return self.get_position(ws).get_cell(ws)


class MarkerPos(MarkerAbs):
    def __init__(self, *args):
        self.pos = Position(*args)

    def get_position(self, ws: Worksheet) -> Position:
        return self.pos


class MarkerName(MarkerAbs):
    def __init__(self, name: str):
        self.name = name

    def get_position(self, ws: Worksheet) -> Position:
        dn = ws.parent.defined_names.get(self.name)
        if dn:
            # Parse defined name
            pos, sheet_name = Position.from_abs_string(dn.value)

            if sheet_name != ws.title:
                raise InvalidPositionError('Marker %s not in worksheet %s' %
                                           (self.name, sheet_name))

            return pos
        else:
            raise InvalidPositionError('Defined name %s not found' % self.name)


class MarkerPattern(MarkerAbs):
    def __init__(self, initial_marker: MarkerAbs, pattern: [str, re.Pattern],
                 direction: Direction, max_range: int):
        self.initial_marker = initial_marker
        self.direction = direction
        self.max_range = max_range
        self.rex = re.compile(pattern)

    def get_position(self, ws: Worksheet) -> Position:
        initial_pos = self.initial_marker.get_position(ws)

        for d in range(self.max_range + 1):
            pos = initial_pos.shifted(self.direction, d)
            val = str(pos.get_cell(ws).value)

            if self.rex.search(val):
                return pos

        raise InvalidPositionError(
            'Cell matching pattern %s max. %d cells %s from %s not found' %
            (str(self.rex), self.max_range, str(
                self.direction), str(initial_pos)))
