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


def get_tables(sheet, table_meta_list):
    row_gen = sheet.get_squared_range(1, 1, sheet.max_column, sheet.max_row)
    cells = [row for row in row_gen]
    tdata = {}
    table_start_end_pairs = _split_table_start_end_pairs(table_meta_list)
    # extract tables
    for tmeta, curr_next_pair in zip(table_meta_list, table_start_end_pairs):
        curr_tmeta, next_tmeta = curr_next_pair
        assert tmeta["name"] == curr_tmeta["name"], \
            "Bad order or curr-next-pairs"
        tdata[tmeta["name"]] = _get_table_data_cells(
            cells, curr_tmeta, next_tmeta)
    return tdata


def _split_table_start_end_pairs(table_meta_list):
    pairs = []
    tno = 0
    while True:
        tmeta = table_meta_list[tno]
        next_tno = tno + 1
        if next_tno == len(table_meta_list):
            pairs.append((tmeta, None))
            break
        else:
            next_tmeta = table_meta_list[next_tno]
            pairs.append((tmeta, next_tmeta))
        tno += 1
    return pairs


def _get_table_data_cells(cells, curr_tmeta, next_tmeta):
    start_found = False
    start = 0
    end_found = False
    end = 0
    if next_tmeta is None:
        end_found = True
        end = len(cells) - 1
    for row, index in zip(cells, range(0, len(cells))):
        c1_value = sanitise_string(row[0].value)
        if not start_found and c1_value.startswith(curr_tmeta["name"]):
            start_found = True
            start = index + 1
        if not end_found and c1_value.startswith(next_tmeta["name"]):
            end_found = True
            end = index
        if start_found and end_found:
            break
    if not start_found or not end_found:
        raise RuntimeError("Could not narrow-down table %s" %
                           (curr_tmeta["name"], ))
    log.debug("Located table %s between rows %d and %d",
              curr_tmeta["name"], start, end)
    table = cells[start:end]
    empty = 0
    for cell in reversed(table[curr_tmeta["first-data-row"] - 1]):
        if cell.value is None:
            empty += 1
        else:
            break
    ntable = []
    log.debug("Trimming extracted table of empty values. Row-len=%d, empty=%d",
              len(row), empty)
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
        return self._variable_names

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
            data = "%s" % (data, )
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


def dump_to_hdf5(data_set, h5_group, dataset_name):
    max_string_length = _compute_max_string_length(data_set)
    dtype = "S%d" % (max_string_length + 1, )
    col_size = len(data_set[0])
    row_size = len(data_set)
    dataset_name = dataset_name.replace('/', ' ')
    h5dset = h5_group.create_dataset(dataset_name, (row_size, col_size),
                                     dtype=dtype)
    #log.debug("Value: %s", data_set)
    for row in range(0, row_size):
        h5dset[row] = data_set[row]


def _compute_max_string_length(data_set):
    col_count = len(data_set[0])
    max_length = 0
    for row in data_set:
        for val in row:
            val_len = len(val)
            if val_len > max_length:
                max_length = val_len
    return max_length
