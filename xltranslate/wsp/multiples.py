#

import logging

import openpyxl

from . import util

log = logging.getLogger(__name__)

TABLES = (
    {
        "name": "Multiples Detail",
        "first-data-row": 2,
    },
)


class TypeBTable(object):
    def __init__(self, table_data):
        self._table_data = table_data
        self._row_len = len(table_data)
        self._col_len = len(table_data[0])
        self._variable_names = []
        self._variable_rows = []
        self._data_set = []
        self._find_variable_info()
        self._extract_data()

    @property
    def variables(self):
        return self._variable_names

    @property
    def data_set(self):
        return self._data_set

    def _find_variable_info(self):
        last_col1_title = None
        for row, index in zip(self._table_data, range(0, self._row_len)):
            col1_data = row[0].value
            if col1_data:
                col1_data = util.sanitise_string(col1_data)
            col2_data = row[1].value
            if col2_data:
                col2_data = util.sanitise_string(col2_data)
            col3_data = row[2].value

            if col3_data is None:
                continue
            if col1_data:
                last_col1_title = col1_data
            var_name = last_col1_title
            if var_name and col2_data:
                var_name = "%s - %s" % (var_name, col2_data)
            if not var_name and col2_data:
                var_name = col2_data
            self._variable_names.append(var_name)
            self._variable_rows.append(index)

    def _extract_data(self):
        self._data_set = []
        for col in range(2, self._col_len):
            self._data_set.append(self._extract_data_for_column(col))

    def _extract_data_for_column(self, col):
        data_set = []
        for row_no in self._variable_rows:
            cell = self._table_data[row_no][col]
            data = cell.value
            if cell.data_type == openpyxl.cell.Cell.TYPE_STRING:
                data = util.sanitise_string(data)
            if cell.is_date:
                data = data.strftime("%b-%d-%Y")
            if data in ('-', ):
                data = 0
            data = str(data)
            data_set.append(data)
        return data_set

    def dump_to_screen(self):
        # construct a format-line for pretty printing
        fmt_list = []
        for col in range(0, len(self._variable_names)):
            col_size = len(self._variable_names[col])
            for row in range(0, len(self._data_set)):
                val = "%s" % (self._data_set[row][col], )
                if val:
                    data_sz = len(val)
                    if data_sz > col_size:
                        col_size = data_sz
            fmt_list.append("{:<%d}" % (col_size, ))
        fmt_line = u' '.join(fmt_list)
        # dump on screen
        print(fmt_line.format(*self._variable_names))
        for ds_line in self._data_set:
            fmtted = fmt_line.format(*ds_line)
            print(fmtted)


class Multiples(object):
    def __init__(self, sheet):
        self._sheet = sheet
        self._raw_tables = util.get_tables(sheet, TABLES)
        self._tables = {}
        for tmeta in TABLES:
            tname = tmeta["name"]
            self._tables[tname] = TypeBTable(self._raw_tables[tname])

    def dump_to_screen(self):
        for tmeta in TABLES:
            tname = tmeta["name"]
            print("\n%s:\n" % (tname, ))
            self._tables[tname].dump_to_screen()

    def dump_to_hdf5(self, h5_group):
        for tmeta in TABLES:
            tname = tmeta["name"]
            table = self._tables[tname]
            util.dump_to_hdf5(table.variables, table.data_set, h5_group, tname)
