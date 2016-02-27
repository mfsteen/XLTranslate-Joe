#

import logging
import argparse

import openpyxl

from . import wsp

log = logging.getLogger(__name__)


def extract(input_file):
    wb = openpyxl.load_workbook(input_file, read_only=True)
    for sheet in wb:
        if sheet.title == "Income Statement":
            wspobj = wsp.IncomeStatement(sheet)
            wspobj.dump()
            continue
        log.warn("Ignoring unknown sheet = %s" % (sheet.title, ))

def main():
    logging.basicConfig(level=logging.DEBUG)

    parser = argparse.ArgumentParser(
        description="Extract data from MS-Excel files")
    parser.add_argument(
        'input_file', metavar="INPUT", type=argparse.FileType('rb'),
        help="File to extract data from")
    args = parser.parse_args()
    extract(args.input_file)
