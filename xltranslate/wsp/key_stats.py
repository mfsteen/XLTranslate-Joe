#
from . import util

TABLES = (
    {
        "name": "Key Financials",
        "first-data-row": 1,
    },
    {
        "name": "Current Capitalization (Millions of USD)",
        "first-data-row": 1,
    },
    {
        "name": "Latest Capitalization (Millions of USD)",
        "first-data-row": 1,
    },
    {
        "name": "Valuation Multiples based on Current Capitalization",
        "first-data-row": 1,
    },
)


def factory(sheet):
    return util.ParsedSheet(sheet, TABLES)
