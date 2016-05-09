#
from . import util

TABLES = (
    {
        "name": "Ratios",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
