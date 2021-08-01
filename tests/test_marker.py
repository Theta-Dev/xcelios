import pytest

# noinspection PyUnresolvedReferences
from tests import workbook, worksheet, worksheet_empty
from xcelios import position


@pytest.mark.parametrize('marker,pos_str', [
    (position.MarkerPos('A3'), 'A3'),
    (position.MarkerName('table_people'), 'B3'),
    (position.MarkerPattern(position.MarkerName('table_people'), r'^Email$',
                            position.Direction.RIGHT, 2), 'D3'),
])
def test_marker(worksheet, marker: position.MarkerAbs, pos_str):
    pos = marker.get_position(worksheet)

    assert str(pos) == pos_str


@pytest.mark.parametrize('marker', [
    position.MarkerName('XYZ'),
    position.MarkerPattern(position.MarkerName('table_people'), r'^XYZ$',
                           position.Direction.RIGHT, 2),
])
def test_marker_err(worksheet, marker: position.MarkerAbs):
    with pytest.raises(position.InvalidPositionError):
        marker.get_position(worksheet)


def test_marker_name_wrong_sheet(worksheet_empty):
    with pytest.raises(position.InvalidPositionError):
        position.MarkerName('table_people').get_position(worksheet_empty)


def test_marker_get_cell(worksheet):
    marker = position.MarkerPos('B4')

    assert marker.get_cell(worksheet).value == 'Hanson'
