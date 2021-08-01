#######
xcelios
#######

Tool for easier data table access with OpenPyXL


Architecture
############

- Objects: Markers, Collections and Datasets

Markers
=======

Markers contain directives that point to a cell of a datasheet.

- Explicit marker (x/y coordinate)
- Named marker (rename a cell in the top left selection field)
- Content marker (content regex + range)

Datasets
========

Every dataset has a cell position as its origin. It can be constructed using a marker pointing to that cell and optionally a collection.

Collections
===========

Collections consist of multiple datasets.

Just like datasets, they have a cell position as their origin.

Collections can provide additional information to the dataset constructor (for example the position of table columns)


Accessing Excel documents
#########################

Reading
=======

- Create a marker pointing to the collection you want to access
- Create the collection:
  - the collection will search for headers and store this information
  - then it will try to construct datasets
- Get your data from the datasets

Writing
=======

- Create a collection
- Add / modify / remove one of its datasets
- The collection will make updated datasets apply their changes (modify/delete associated cells) or write a new dataset into the table

Formatting
==========

Datasets and collections can have a format method to apply formatting to the associated cells. Formatting needs to be applied to newly created datasets and collections.

How to shrink/extend tables
###########################

- Tables have a minimum distance to other content (by default 2 cells)
- Before data is written back to the table, the available space has to be determined.
- Is there less space than required by the datasets? Then the space needs to be extended
  - Starting from the bottom (the row before the following content) moving upward, find a row that is empty except for the table's own data
  - Insert the required amount of new rows there
- Is there more space than required? Try to contract the table
  - Starting from the bottom moving upward, look for rows that are empty except for the table's own data
  - Delete these rows until the required space is reached OR
  - If a row containing data is encountered, stop to avoid destroying the table's layout.


What currently works
====================

.. code-block:: python

  from xcelios import table, position
  from openpyxl import open
  from dataclasses import dataclass

  # Define a data class
  # By default, variable names correspond to the table header names
  # Underscores translate to spaces, dashes, underscores or nothing
  # in the header name

  # Possible extension: Allow names with special characters
  # as header names using docstrings or additional methods
  @dataclass
  class Person:
      first_name: str
      last_name: str
      email: str
      birthday: str
      height: int
      favorite_food: str

  # Open workbook and worksheet
  wb = open('./tests/testfiles/Test1.xlsx')
  ws = wb['Sheet1']

  # Construct a table with the worksheet, a marker pointing to a corner point
  # and the data class
  tab = table.Table(ws, position.MarkerName('table_people'), Person)

  # Read and access the data
  tab.read_datasets()

  tab.datasets
  # [Person(first_name='Hanson', last_name='Marnane', email='hmarnane0@arizona.edu', ...

  # Write back data after modification
  tab.write_datasets()

  # Save the modified worksheet
  wb.save('output.xlsx')
