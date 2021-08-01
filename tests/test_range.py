import pytest

from xcelios.position import (Direction, InvalidPositionError,
                              InvalidRangeError, Position, Range)


@pytest.mark.parametrize('pos_a,pos_b,range_str', [
    ('A1', 'C3', 'A1:C3'),
    ('C3', 'A1', 'A1:C3'),
])
def test_from_pos(pos_a, pos_b, range_str):
    rg = Range.from_pos(Position(pos_a), Position(pos_b))

    assert str(rg) == range_str


@pytest.mark.parametrize('range_str', [
    'A1:D10',
    'D10:XFD1048576',
])
def test_from_str(range_str):
    assert str(Range.from_str(range_str)) == range_str


@pytest.mark.parametrize('mr,xr,mc,xc,exception', [
    (10, 5, 1, 2, InvalidRangeError),
    (1, 2, 10, 5, InvalidRangeError),
    (-1, 2, 5, 10, InvalidPositionError),
])
def test_range_err(mr, xr, mc, xc, exception):
    with pytest.raises(exception):
        Range(mr, xr, mc, xc)


@pytest.mark.parametrize('rg,pos,is_in', [
    (Range.from_str('B2:D10'), Position('B6'), True),
    (Range.from_str('B2:D10'), Position('A1'), False),
])
def test_is_inside(rg, pos, is_in):
    assert rg.is_inside(pos) == is_in


@pytest.mark.parametrize('rg,direction,distance,n_rg', [
    (Range.from_str('B2:D10'), Direction.RIGHT, 1, Range.from_str('B2:E10')),
    (Range.from_str('B2:D10'), Direction.DOWN, 1, Range.from_str('B2:D11')),
])
def test_extended(rg, direction, distance, n_rg):
    assert rg.extended(direction, distance) == n_rg
