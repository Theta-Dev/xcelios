import os
from dataclasses import dataclass
from datetime import datetime

import pytest

# noinspection PyUnresolvedReferences
from tests import DIR_JSON, assert_obj_equals_json_file, workbook, worksheet
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
    sum: str


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


@dataclass
class PersonTypeErr:
    first_name: int
    last_name: str
    email: str
    birthday: datetime
    height: int
    favorite_food: str


@pytest.mark.parametrize('marker_name,args,header_pos', [
    ('table_people', [Person], {
        'first_name': position.Position('B3'),
        'last_name': position.Position('C3'),
        'email': position.Position('D3'),
        'birthday': position.Position('E3'),
        'height': position.Position('F3'),
        'favorite_food': position.Position('G3'),
    }),
    ('table_prices',
     [Prices, position.Direction.DOWN, position.Direction.RIGHT], {
         'date': position.Position('B24'),
         'product_a': position.Position('B26'),
         'product_b': position.Position('B27'),
         'sum': position.Position('B28'),
     }),
])
def test_table_headers(worksheet, marker_name, args, header_pos):
    tb = table.Table(worksheet, position.MarkerName(marker_name), *args)

    assert tb.title_positions == header_pos


@pytest.mark.parametrize('marker_name,args,emsg', [
    ('table_people', [PersonErr], 'Could not find table headers: lol, wtf'),
    ('table_prices', [
        Person, position.Direction.DOWN, position.Direction.RIGHT
    ], 'Could not find table headers: first_name, last_name, email, birthday, \
height, favorite_food'),
    ('table_prices', [
        Prices, position.Direction.DOWN, position.Direction.RIGHT, 0
    ], 'Could not find table headers: product_a, product_b, sum'),
])
def test_table_headers_err(worksheet, marker_name, args, emsg):
    with pytest.raises(table.TableParseError) as e:
        table.Table(worksheet, position.MarkerName(marker_name), *args)

    assert str(e.value) == emsg


@pytest.mark.parametrize('marker_name,args,json_file', [
    ('table_people', [Person], 'people.json'),
    ('table_prices', [
        Prices, position.Direction.DOWN, position.Direction.RIGHT
    ], 'prices.json'),
])
def test_read_datasets(worksheet, marker_name, args, json_file):
    tab = table.Table(worksheet, position.MarkerName(marker_name), *args)
    tab.read_datasets()

    assert_obj_equals_json_file(tab.datasets,
                                os.path.join(DIR_JSON, json_file))
