#
from . import util

TABLES = (
    {
        "name": "Cash Flow",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
