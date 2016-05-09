#
from . import util

TABLES = (
  {
      "name": "Pension/OPEB",
      "first-data-row": 1,
  },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
