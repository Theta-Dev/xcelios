from dataclasses import dataclass
from datetime import datetime

import openpyxl
import pytest

import tests
from xcelios import position, table


@dataclass
class Person:
    first_name: str
    last_name: str
    email: str
    birthday: datetime
    height: int
    favorite_food: str


@dataclass
class Prices:
    date: datetime
    product_a: float
    product_b: float
    sum: float


@dataclass
class PersonErr:
    first_name: str
    last_name: str
    email: str
    birthday: datetime
    height: int
    favorite_food: str
    lol: str
    wtf: str


@pytest.mark.parametrize('marker_name,args,header_pos', [
    ('table_people', [Person], {
        'first_name': position.Position('Sheet1', 'B3'),
        'last_name': position.Position('Sheet1', 'C3'),
        'email': position.Position('Sheet1', 'D3'),
        'birthday': position.Position('Sheet1', 'E3'),
        'height': position.Position('Sheet1', 'F3'),
        'favorite_food': position.Position('Sheet1', 'G3'),
    }),
    ('table_prices',
     [Prices, position.Direction.DOWN, position.Direction.RIGHT], {
         'date': position.Position('Sheet1', 'B24'),
         'product_a': position.Position('Sheet1', 'B25'),
         'product_b': position.Position('Sheet1', 'B26'),
         'sum': position.Position('Sheet1', 'B27'),
     }),
])
def test_table_headers(marker_name, args, header_pos):
    wb = openpyxl.open(tests.FILE_TEST1)
    tb = table.Table(wb, position.MarkerName(marker_name), *args)

    assert tb.title_positions == header_pos

    wb.close()


@pytest.mark.parametrize('marker_name,args,emsg', [
    ('table_people', [PersonErr], 'Could not find table headers: lol, wtf'),
    ('table_prices', [
        Person, position.Direction.DOWN, position.Direction.RIGHT
    ], 'Could not find table headers: first_name, last_name, email, birthday, \
height, favorite_food'),
])
def test_table_headers_err(marker_name, args, emsg):
    wb = openpyxl.open(tests.FILE_TEST1)

    with pytest.raises(table.TableParseError) as e:
        table.Table(wb, position.MarkerName(marker_name), *args)

    assert str(e.value) == emsg

    wb.close()
