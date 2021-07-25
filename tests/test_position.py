import openpyxl
import pytest

import tests
from xcelios import position


@pytest.mark.parametrize('args,pos_str', [
    (['Sheet1', 'A', 1], 'Sheet1!$A$1'),
    (['Sheet1', 2, 3], 'Sheet1!$B$3'),
    (['Sheet1', 'C', '6'], 'Sheet1!$C$6'),
    (['Sheet1', 'A1'], 'Sheet1!$A$1'),
    (['Sheet1', 'AB12'], 'Sheet1!$AB$12'),
    (['Sheet1!$AB$12'], 'Sheet1!$AB$12'),
    (['Sheet1', 'XFD', 1048576], 'Sheet1!$XFD$1048576'),
])
def test_position_obj(args, pos_str):
    pos = position.Position(*args)

    assert str(pos) == pos_str
    assert repr(pos) == '<Position: %s>' % pos_str


@pytest.mark.parametrize('args', [
    ['Sheet1', 'XFE', 1048576],
    ['Sheet1', 'A', 1048577],
])
def test_position_obj_err(args):
    with pytest.raises(position.InvalidPositionError):
        position.Position(*args)


def test_position_obj_terr():
    with pytest.raises(TypeError):
        position.Position()

    with pytest.raises(TypeError):
        position.Position('Sheet1', 'A', 1, 'X')


@pytest.mark.parametrize('marker,pos_str', [
    (position.MarkerPos('Sheet1', 'A3'), 'Sheet1!$A$3'),
    (position.MarkerName('table_people'), 'Sheet1!$B$3'),
    (position.MarkerPattern(position.MarkerName('table_people'), r'^Email$',
                            position.Direction.RIGHT, 2), 'Sheet1!$D$3'),
])
def test_marker(marker: position.MarkerAbs, pos_str):
    wb = openpyxl.open(tests.FILE_TEST1)
    pos = marker.get_position(wb)

    assert str(pos) == pos_str
    wb.close()


@pytest.mark.parametrize('marker', [
    position.MarkerPos('SheetX', 'A3'),
    position.MarkerName('XYZ'),
    position.MarkerPattern(position.MarkerName('table_people'), r'^XYZ$',
                           position.Direction.RIGHT, 2),
])
def test_marker_err(marker: position.MarkerAbs):
    wb = openpyxl.open(tests.FILE_TEST1)

    with pytest.raises(position.InvalidPositionError):
        marker.get_position(wb)

    wb.close()
