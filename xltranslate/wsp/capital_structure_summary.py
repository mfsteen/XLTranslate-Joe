#
import logging

import openpyxl

from . import util

log = logging.getLogger(__name__)

TABLES = (
    {
        "name": "Capital Structure Data",
        "first-data-row": 3,
    },
    {
        "name": "Debt Summary Data",
        "first-data-row": 3,
    },
)


def _trim_table(table):
    new_table = []
    for row in table:
        col_data = row[0].value
        if col_data is None:
            continue
        col_data = row[1].value
        if col_data is None:
            continue
        new_table.append(row)
    return new_table


def _fiscal_period_start_columns(table):
    row = table[0]
    columns = []
    for col in range(1, len(row), 2):
        data = row[col].value
        if data is None:
            break
        columns.append(col)
    return columns


def _variables_meta(table):
    variables = []
    locations = []
    variables.append(util.sanitise_string(table[0][0].value))
    locations.append((0, 0))
    variables.append(util.sanitise_string(table[1][0].value))
    locations.append((1, 0))
    type1 = util.sanitise_string(table[2][1].value)
    type2 = util.sanitise_string(table[2][2].value)
    for row_no in range(3, len(table)):
        var_name = util.sanitise_string(table[row_no][0].value)
        var1_name = "%s - %s" % (var_name, type1)
        var2_name = "%s - %s" % (var_name, type2)
        variables.append(var1_name)
        locations.append((row_no, 0))
        variables.append(var2_name)
        locations.append((row_no, 1))
    return variables, locations


def _extract_from_col(table, locations, start_col):
    drow = []
    for row_no, col_offset in locations:
        cell = table[row_no][start_col + col_offset]
        data = cell.value
        if cell.data_type == openpyxl.cell.Cell.TYPE_STRING:
            data = util.sanitise_string(data)
        if cell.is_date:
            data = data.strftime("%b-%d-%Y")
        if data in ('-', ):
            data = 0
        data = str(data)
        drow.append(data)
    return drow


def _extract(table):
    table = _trim_table(table)
    fiscal_period_cols = _fiscal_period_start_columns(table)
    if len(fiscal_period_cols) == 0:
        log.error("No fiscal data columns. Bailing.")
        return None, None
    variables, locations = _variables_meta(table)
    data_set = []
    for start_col in fiscal_period_cols:
        row = _extract_from_col(table, locations, start_col)
        data_set.append(row)
    return variables, data_set


class Table(object):
    def __init__(self, variables, data_set):
        self._variables = variables
        self._data_set = data_set

    @property
    def variables(self):
        return self._variables

    @property
    def data_set(self):
        return self._data_set

    def dump(self):
        # construct a format-line for pretty printing
        fmt_list = []
        for col in range(0, len(self._variables)):
            col_size = len(self._variables[col])
            for row in range(0, len(self._data_set)):
                val = "%s" % (self._data_set[row][col], )
                if val:
                    data_sz = len(val)
                    if data_sz > col_size:
                        col_size = data_sz
            fmt_list.append("{:<%d}" % (col_size, ))
        fmt_line = u' '.join(fmt_list)
        # dump on screen
        print(fmt_line.format(*self._variables))
        for ds_line in self._data_set:
            fmtted = fmt_line.format(*ds_line)
            print(fmtted)


def extract_type_c_table(table):
    variables, data_set = _extract(table)
    if variables is None:
        return None
    table = Table(variables, data_set)
    return table


class CapitalStructureSummary(object):
    def __init__(self, sheet):
        raw_tables = util.get_tables(sheet, TABLES)
        self._tables = {}
        for tmeta in TABLES:
            tname = tmeta["name"]
            table = extract_type_c_table(raw_tables[tname])
            if table is None:
                log.error("Failed to extract table '%s'", tname)
            self._tables[tname] = table

    def dump_to_screen(self):
        for name, table in self._tables.iteritems():
            print("\n%s:\n" % (name, ))
            table.dump_to_screen()

    def dump_to_hdf5(self, h5_group):
        for tmeta in TABLES:
            tname = tmeta["name"]
            table = self._tables[tname]
            util.dump_to_hdf5(table.variables, table.data_set, h5_group, tname)
