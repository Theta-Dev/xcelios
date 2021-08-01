import datetime
import json
import os
from typing import Callable
from unittest.mock import ANY

import openpyxl
import pytest
from importlib_resources import files
from openpyxl.worksheet.worksheet import Worksheet

DIR_TESTFILES = str(files('tests').joinpath('testfiles'))
DIR_JSON = os.path.join(DIR_TESTFILES, 'json')
FILE_TEST1 = os.path.join(DIR_TESTFILES, 'Test1.xlsx')


@pytest.fixture(scope='module')
def workbook() -> openpyxl.Workbook:
    wb = openpyxl.open(FILE_TEST1)
    yield wb
    wb.close()


@pytest.fixture(scope='module')
def worksheet(workbook) -> Worksheet:
    return workbook['Sheet1']


@pytest.fixture(scope='module')
def worksheet_empty(workbook) -> Worksheet:
    return workbook['SheetE']


def get_np_attrs(o) -> dict:
    """
    Return all non-protected attributes of the given object.

    :param o:  Object
    :return: Dict of attributes
    """
    return {
        key: getattr(o, key)
        for key in o.__dir__() if not key.startswith('_')
        and not isinstance(getattr(o, key), Callable)
    }


def _serializer(o):
    if hasattr(o, 'serialize'):
        return o.serialize()
    elif isinstance(o, datetime.datetime) or isinstance(o, datetime.date):
        return o.isoformat()
    elif hasattr(o, '__dict__'):
        return get_np_attrs(o)
    return str(o)


def to_json(o, pretty=True) -> str:
    """
    Convert object to json.
    Uses the ``serialize()`` method of the target object if available.

    :param o: Object to serialize
    :param pretty: Prettify with indents
    :return: JSON string
    """
    return json.dumps(o,
                      default=_serializer,
                      indent=2 if pretty else None,
                      ensure_ascii=False)


def to_json_file(o, path):
    """
    Convert object to json and writes the result to a file.
    Uses the ``serialize()`` method of the target object if available.

    :param o: Object to serialize
    :param path: File path
    """
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(o, f, default=_serializer, indent=2, ensure_ascii=False)


def _rec_iterate(o, fun):
    """
    Iterate through common data structures consisting of dicts and lists,
    modifying contained objects using the given function.

    :param o: Data structure
    :param fun: Modifier function: fun(obj) -> new_ob
    :return: New data structure
    """
    if isinstance(o, dict):
        return {k: _rec_iterate(x, fun) for k, x in o.items()}
    if isinstance(o, list):
        return [_rec_iterate(x, fun) for x in o]
    return fun(o)


def _replace_any(o):
    """
    Replace the string ``<<ANY>>`` with mock.ANY.
    Used together with rec_iterate
    """
    if o == '<<ANY>>':
        return ANY
    return o


def assert_obj_equals_json_file(o, file, writeback=False):
    """
    Serialize an object and verify if it matches the given JSON file.

    If you enable the ``writeback`` option, it will create a new json file named
    <FILENAME>wb, which contains the correct representation of the object
    (only if it differs from the content of the original json file).

    :param o: Serializable object
    :param file: Path to JSON file
    :param writeback: Enable writeback option
    """
    with open(file, 'r', encoding='utf-8') as f:
        file_data = json.load(f)

    exp_data = _rec_iterate(file_data, _replace_any)

    obj_json = to_json(o)
    obj_data = json.loads(obj_json)

    print('---BEGIN JSON DATA---')
    print(obj_json)
    print('---END JSON DATA---')

    # Optionally write to file for comparison
    if writeback and obj_data != exp_data:
        dir_path, basename = os.path.split(file)
        n_file = os.path.join(dir_path,
                              os.path.splitext(basename)[0] + '.wb.json')

        with open(n_file, 'w', encoding='utf-8') as f:
            f.write(obj_json)

    assert obj_data == exp_data
