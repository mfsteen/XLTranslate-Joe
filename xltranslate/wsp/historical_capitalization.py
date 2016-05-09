#
from . import util

TABLES = (
    {
        "name": "Historical Capitalization",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
