from typing import Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.cell import column_index_from_string, get_column_letter


class Position:
    def __init__(self, ws: str, row: [int, str], col: [int, str]):
        pass


class Marker:
    def __init__(self, parent):
        pass

    def get_position(self) -> Position:
        pass
