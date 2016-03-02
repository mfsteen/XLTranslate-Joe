#

import logging

import openpyxl

log = logging.getLogger(__name__)


def sanitise_string(text):
    text = u'%s' % (text, )
    text = text.strip()
    text = text.replace('\n', ' ')
    text = text.encode('ascii', 'ignore')
    return text


def get_tables(sheet, table_name_list):
    row_gen = sheet.get_squared_range(1, 1, sheet.max_column, sheet.max_row)
    cells = [row for row in row_gen]
    tdata = {}
    table_start_end_pairs = _split_table_start_end_pairs(table_name_list)
    # extract tables
    for tname, curr_next_pair in zip(table_name_list, table_start_end_pairs):
        curr_tname, next_tname = curr_next_pair
        assert tname == curr_tname, "Bad order or curr-next-pairs"
        tdata[tname] = _get_table_data_cells(cells, curr_tname, next_tname)
    return tdata


def _split_table_start_end_pairs(table_name_list):
    pairs = []
    tno = 0
    while True:
        tname = table_name_list[tno]
        next_tno = tno + 1
        if next_tno == len(table_name_list):
            pairs.append((tname, None))
            break
        else:
            next_tname = table_name_list[next_tno]
            pairs.append((tname, next_tname))
        tno += 1
    return pairs


def _get_table_data_cells(cells, curr_table_name, next_table_name):
    start_found = False
    start = 0
    end_found = False
    end = 0
    if next_table_name is None:
        end_found = True
        end = len(cells) - 1
    for row, index in zip(cells, range(0, len(cells))):
        c1_value = sanitise_string(row[0].value)
        if not start_found and c1_value == curr_table_name:
            start_found = True
            start = index + 1
        if not end_found and c1_value == next_table_name:
            end_found = True
            end = index
        if start_found and end_found:
            break
    if not start_found or not end_found:
        raise RuntimeError("Could not narrow-down table %s" %
                           (curr_table_name, ))
    table = cells[start:end]
    empty = 0
    for cell in reversed(table[0]):
        if cell.value is None:
            empty += 1
        else:
            break
    log.debug("found %d empty cells in table %s", empty, curr_table_name)
    ntable = []
    for row in table:
        ntable.append(row[0:(len(row) - empty)])
    return ntable


class TypeATable(object):
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
        return self._variables

    @property
    def data_set(self):
        return self._data_set

    def _find_variable_info(self):
        for row, index in zip(self._table_data, range(0, self._row_len)):
            col1_data = sanitise_string(row[0].value)
            col2_data = row[1].value
            if col1_data and col2_data:
                self._variable_names.append(col1_data)
                self._variable_rows.append(index)

    def _extract_data(self):
        self._data_set = []
        for col in range(1, self._col_len):
            self._data_set.append(self._extract_data_for_column(col))

    def _extract_data_for_column(self, col):
        data_set = []
        for row_no in self._variable_rows:
            cell = self._table_data[row_no][col]
            data = cell.value
            if cell.data_type == openpyxl.cell.Cell.TYPE_STRING:
                data = sanitise_string(data)
            if cell.is_date:
                data = data.strftime("%b-%d-%Y")
            if data in ('-', ):
                data = 0
            data_set.append(data)
        return data_set

    def dump(self):
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
