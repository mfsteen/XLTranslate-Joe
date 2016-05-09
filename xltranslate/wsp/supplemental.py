#
from . import util

TABLES = (
    {
        "name": "Supplemental",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
