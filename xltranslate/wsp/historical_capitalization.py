#
import logging

from . import util

log = logging.getLogger(__name__)

TABLES = (
    {
        "name": "Historical Capitalization",
        "first-data-row": 1,
    },
)


class HistoricalCapitalization(object):
    def __init__(self, sheet):
        self._sheet = sheet
        self._raw_tables = util.get_tables(sheet, TABLES)
        self._tables = {}
        for tmeta in TABLES:
            tname = tmeta["name"]
            self._tables[tname] = util.TypeATable(self._raw_tables[tname])

    def dump(self):
        for tmeta in TABLES:
            tname = tmeta["name"]
            print("\n%s:\n" % (tname, ))
            self._tables[tname].dump()
