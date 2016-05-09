#

from . import util

TABLES = (
    {
        "name": "Balance Sheet",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
