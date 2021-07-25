import os

from importlib_resources import files

DIR_TESTFILES = str(files('tests.testfiles').joinpath(''))
FILE_TEST1 = os.path.join(DIR_TESTFILES, 'Test1.xlsx')
