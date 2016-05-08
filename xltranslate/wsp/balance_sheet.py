#
import logging

from . import util

log = logging.getLogger(__name__)

TABLES = (
    {
        "name": "Balance Sheet",
        "first-data-row": 1,
    },
)


class BalanceSheet(object):
    def __init__(self, sheet):
        self._sheet = sheet
        self._metadata = util.get_sheet_metadata(sheet)
        self._raw_tables = util.get_tables(sheet, TABLES)
        self._tables = {}
        for tmeta in TABLES:
            tname = tmeta["name"]
            self._tables[tname] = util.TypeATable(self._raw_tables[tname])

    def dump_to_screen(self):
        for tmeta in TABLES:
            tname = tmeta["name"]
            print("\n%s:\n" % (tname, ))
            self._tables[tname].dump_to_screen()

    def dump_to_hdf5(self, h5_group):
        # Store sheet metadata as group attributes
        for k, v in self._metadata.items():
            h5_group.attrs[k] = v
        # Store each table as dataset
        for tmeta in TABLES:
            tname = tmeta["name"]
            table = self._tables[tname]
            util.dump_to_hdf5(table.variables, table.data_set, h5_group, tname)
