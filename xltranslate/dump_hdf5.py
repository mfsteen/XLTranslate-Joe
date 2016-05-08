# To dump a hdf5 file

import logging
import argparse

import h5py

log = logging.getLogger(__name__)


def _get_fmt_string(data_set):
    max_len = 0
    for row in data_set:
        for val in row:
            if len(val) > max_len:
                max_len = len(val)
    fmt_string_list = []
    row_len, col_len = data_set.shape
    for index in range(0, col_len):
        fmt_string_list.append("{:<%d}" % (max_len, ))
    return u' '.join(fmt_string_list)


def dump(file_name):
    h5 = h5py.File(file_name, 'r')
    for group_name in h5:
        group = h5[group_name]
        for dataset_name in group:
            dataset = group[dataset_name]
            log.debug("Group: %s, Dataset: %s", group_name, dataset_name)
            fmt_string = _get_fmt_string(dataset)
            for row in dataset:
                print(fmt_string.format(*row))


def main():
    logging.basicConfig(level=logging.DEBUG)

    parser = argparse.ArgumentParser(
        description="Dump data from a given HDF5 file")
    parser.add_argument(
        'input_file', metavar="INPUT", type=argparse.FileType('rb'),
        help="File to read data from")
    args = parser.parse_args()
    dump(args.input_file.name)
