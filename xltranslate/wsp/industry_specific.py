#
import logging

from . import util

log = logging.getLogger(__name__)

TABLES = (
    {
        "name": "Industry Specific",
        "first-data-row": 1,
    },
)


class IndustrySpecific(object):
    def __init__(self, sheet):
        self._sheet = sheet
        self._metadata = util.get_sheet_metadata(sheet)
        self._raw_tables = util.get_tables(sheet, TABLES)
        self._tables = {}
        for tmeta in TABLES:
            tname = tmeta["name"]
            table = util.create_type_a_table(self._raw_tables[tname])
            if table is None:
                log.info("In sheet '%s', ignoring empty table: '%s'",
                         sheet.title, tname)
                continue
            self._tables[tname] = table

    @property
    def tables(self):
        return self._tables

    @property
    def metadata(self):
        return self._metadata

    def dump_to_screen(self):
        for tmeta in TABLES:
            tname = tmeta["name"]
            print("\n%s:\n" % (tname, ))
            self._tables[tname].dump_to_screen()
