import openpyxl
import pytest

import tests
from xcelios import position


@pytest.mark.parametrize('args,pos_str', [
    (['A', 1], 'A1'),
    ([2, 3], 'B3'),
    (['C', '6'], 'C6'),
    (['A1'], 'A1'),
    (['AB12'], 'AB12'),
    (['XFD', 1048576], 'XFD1048576'),
])
def test_position_obj(args, pos_str):
    pos = position.Position(*args)

    assert str(pos) == pos_str
    assert repr(pos) == '<Position: %s>' % pos_str


@pytest.mark.parametrize('args', [
    ['XFE', 1048576],
    ['A', 1048577],
    ['F'],
    ['#', 1],
    ['A', 'A'],
])
def test_position_obj_err(args):
    with pytest.raises(position.InvalidPositionError):
        position.Position(*args)


def test_position_obj_terr():
    with pytest.raises(TypeError):
        position.Position()

    with pytest.raises(TypeError):
        position.Position('A', 1, 'X')


@pytest.mark.parametrize('abs_str,pos,sheet_name', [
    ('Sheet1!$A$4', position.Position('A', 4), 'Sheet1'),
    ('SheetX!$ZA$115', position.Position('ZA', 115), 'SheetX'),
])
def test_position_from_abs(abs_str, pos, sheet_name):
    p, n = position.Position.from_abs_string(abs_str)

    assert p == pos
    assert n == sheet_name


@pytest.mark.parametrize('abs_str', [
    'XYZ',
    'Sheet1!$A$B',
])
def test_position_from_abs_err(abs_str):
    with pytest.raises(position.InvalidPositionError):
        position.Position.from_abs_string(abs_str)


@pytest.mark.parametrize('marker,pos_str', [
    (position.MarkerPos('A3'), 'A3'),
    (position.MarkerName('table_people'), 'B3'),
    (position.MarkerPattern(position.MarkerName('table_people'), r'^Email$',
                            position.Direction.RIGHT, 2), 'D3'),
])
def test_marker(marker: position.MarkerAbs, pos_str):
    wb = openpyxl.open(tests.FILE_TEST1)
    ws = wb['Sheet1']
    pos = marker.get_position(ws)

    assert str(pos) == pos_str
    wb.close()


@pytest.mark.parametrize('marker', [
    position.MarkerName('XYZ'),
    position.MarkerPattern(position.MarkerName('table_people'), r'^XYZ$',
                           position.Direction.RIGHT, 2),
])
def test_marker_err(marker: position.MarkerAbs):
    wb = openpyxl.open(tests.FILE_TEST1)
    ws = wb['Sheet1']

    with pytest.raises(position.InvalidPositionError):
        marker.get_position(ws)

    wb.close()


def test_marker_name_wrong_sheet():
    wb = openpyxl.open(tests.FILE_TEST1)
    ws = wb['SheetE']

    with pytest.raises(position.InvalidPositionError):
        position.MarkerName('table_people').get_position(ws)

    wb.close()


def test_marker_get_cell():
    wb = openpyxl.open(tests.FILE_TEST1)
    ws = wb['Sheet1']
    marker = position.MarkerPos('B4')

    assert marker.get_cell(ws).value == 'Hanson'

    wb.close()
