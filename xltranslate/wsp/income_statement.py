#
from . import util

TABLES = (
    {
        "name": "Income Statement",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
