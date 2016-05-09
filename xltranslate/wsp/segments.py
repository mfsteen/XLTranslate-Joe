#
from . import util

TABLES = (
    {
        "name": "Business Segments",
        "first-data-row": 1,
    },
    {
        "name": "Geographic Segments",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
