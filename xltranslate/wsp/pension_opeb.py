#
import logging

from . import util

log = logging.getLogger(__name__)

TABLES = (
    "Pension/OPEB",
)


class PensionOPEB(object):
    def __init__(self, sheet):
        self._sheet = sheet
        self._raw_tables = util.get_tables(sheet, TABLES)
        self._tables = {}
        for tname in TABLES:
            self._tables[tname] = util.TypeATable(self._raw_tables[tname])

    def dump(self):
        for tname in TABLES:
            print("\n%s:\n" % (tname, ))
            self._tables[tname].dump()
