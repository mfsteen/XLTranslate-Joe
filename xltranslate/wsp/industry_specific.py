#
from . import util

TABLES = (
    {
        "name": "Industry Specific",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
