import re
from typing import Any, Type

from openpyxl import Workbook

from xcelios.position import Direction, MarkerAbs


class TableParseError(Exception):
    pass


class Dataset:
    def __init__(self, line: int, value: Any):
        self.line: line
        self.value = value


class Table:
    def __init__(self,
                 wb: Workbook,
                 initial_marker: MarkerAbs,
                 obj_class: Type,
                 header_dir: Direction = Direction.RIGHT,
                 body_dir: Direction = Direction.DOWN):
        self.wb = wb
        self.initial_pos = initial_marker.get_position(self.wb)
        self.obj_class = obj_class
        self.header_dir = header_dir
        self.body_dir = body_dir

        # Maximum number of blank rows/cols to ignore
        self.max_blanks = 1

        self.title_positions = dict()
        self.datasets = []

        self._locate_headers()

    def _locate_headers(self):
        # Read type annotations from dataclass
        title_rexes = dict()
        for key in self.obj_class.__annotations__.keys():
            title_rexes[key] = re.compile(
                re.escape(key).replace('_', r'[_\- ]?'), re.IGNORECASE)

        blanks = 0
        pos = self.initial_pos

        # Stop iteration after encountering more than max_blanks empty cells
        # after eachother, reaching the end of the worksheet
        # or having found all titles
        while blanks <= self.max_blanks and pos.is_in(self.wb) and title_rexes:
            val = str(pos.get_cell(self.wb).value)

            if val:
                blanks = 0
                found_key = None

                for key, rex in title_rexes.items():
                    if rex.match(val):
                        found_key = key
                        break

                if found_key is not None:
                    self.title_positions[found_key] = pos
                    title_rexes.pop(found_key)
            else:
                blanks += 1

            pos = pos.shifted(self.header_dir)

        if title_rexes:
            raise TableParseError('Could not find table headers: %s' %
                                  ', '.join(title_rexes.keys()))

    def read_datasets(self):
        self.datasets = []
