import pytest

# noinspection PyUnresolvedReferences
from tests import workbook, worksheet, worksheet_empty
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


def test_position_eq():
    assert position.Position('A1') == position.Position('A1')
    assert position.Position('A1') != position.Position('A2')


@pytest.mark.parametrize('origin,direction,d,target', [
    ('A1', position.Direction.DOWN, 1, 'A2'),
    ('C3', position.Direction.UP, 2, 'C1'),
    ('C3', position.Direction.RIGHT, 4, 'G3'),
    ('C3', position.Direction.LEFT, 2, 'A3'),
])
def test_shifted(origin, direction, d, target):
    pos = position.Position(origin)
    shift = pos.shifted(direction, d)

    assert str(shift) == target


@pytest.mark.parametrize('pos_a,pos_b,direction,distance', [
    ('A1', 'A10', position.Direction.DOWN, 9),
    ('A10', 'A1', position.Direction.DOWN, -9),
    ('A10', 'A1', position.Direction.UP, 9),
    ('A1', 'F1', position.Direction.RIGHT, 5),
    ('F1', 'A1', position.Direction.RIGHT, -5),
    ('F1', 'A1', position.Direction.LEFT, 5),
])
def test_dir_distance(pos_a, pos_b, direction, distance):
    pa = position.Position(pos_a)
    pb = position.Position(pos_b)

    dst = pa.dir_distance(pb, direction)

    assert dst == distance
    assert pa.shifted(direction, dst) == pb


@pytest.mark.parametrize('pos_a,pos_b,axis,pos_c', [
    ('A1', 'Z12', position.Axis.ROW, 'Z1'),
    ('A1', 'Z12', position.Axis.COL, 'A12'),
])
def test_combine(pos_a, pos_b, axis, pos_c):
    pa = position.Position(pos_a)
    pb = position.Position(pos_b)

    comb = pa.combine(axis, pb)

    assert str(comb) == pos_c


@pytest.mark.parametrize('pos_a,n_b,axis,pos_c', [
    ('A1', 26, position.Axis.ROW, 'Z1'),
    ('A1', 12, position.Axis.COL, 'A12'),
])
def test_combine_int(pos_a, n_b, axis, pos_c):
    pa = position.Position(pos_a)

    comb = pa.combine(axis, n_b)

    assert str(comb) == pos_c


@pytest.mark.parametrize('sname,pos,is_in', [
    ('Sheet1', 'A1', False),
    ('Sheet1', 'B3', True),
    ('Sheet1', 'B29', False),
    ('SheetE', 'A1', True),
    ('SheetE', 'B3', False),
])
def test_is_in(workbook, sname, pos, is_in):
    ws = workbook[sname]
    pos = position.Position(pos)

    assert pos.is_in(ws) == is_in


def test_get_cell(worksheet):
    pos = position.Position('B4')

    assert pos.get_cell(worksheet).value == 'Hanson'


@pytest.mark.parametrize('pos,empty', [
    ('A1', True),
    ('B3', False),
])
def test_is_cell_empty(worksheet, pos, empty):
    pos = position.Position(pos)

    assert pos.is_cell_empty(worksheet) == empty


@pytest.mark.parametrize('pos,axis,coord', [
    ('A10', position.Axis.ROW, 10),
    ('A10', position.Axis.COL, 1),
])
def test_get_coord(pos, axis, coord):
    assert position.Position(pos).get_coord(axis) == coord
