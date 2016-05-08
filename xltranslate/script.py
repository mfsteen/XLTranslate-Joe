#

import logging
import argparse

import openpyxl
import h5py

from . import wsp

log = logging.getLogger(__name__)

WSP_REGISTRY = {
    'Key Stats': wsp.KeyStats,
    'Income Statement': wsp.IncomeStatement,
    'Balance Sheet': wsp.BalanceSheet,
    'Cash Flow': wsp.CashFlow,
    "Multiples": wsp.Multiples,
    "Historical Capitalization": wsp.HistoricalCapitalization,
    "Capital Structure Summary": wsp.CapitalStructureSummary,
    "Ratios": wsp.Ratios,
    "Supplemental": wsp.Supplemental,
    "Industry Specific": wsp.IndustrySpecific,
    "Pension OPEB": wsp.PensionOPEB,
    "Segments": wsp.Segments,
}


def extract(input_file, output_file_name):
    wb = openpyxl.load_workbook(input_file, read_only=True)
    h5 = h5py.File(output_file_name, "w")
    for sheet in wb:
        klass = WSP_REGISTRY.get(sheet.title, None)
        if klass is not None:
            wspobj = klass(sheet)
            group = h5.create_group(sheet.title)
            wspobj.dump_to_hdf5(group)
        else:
            log.warn("Ignoring unknown sheet = %s" % (sheet.title, ))


def main():
    logging.basicConfig(level=logging.DEBUG)

    parser = argparse.ArgumentParser(
        description="Extract data from MS-Excel files")
    parser.add_argument(
        'input_file', metavar="INPUT", type=argparse.FileType('rb'),
        help="File to extract data from")
    parser.add_argument(
        'output_file', metavar="OUTPUT", type=argparse.FileType('w'),
        help="File to store extracted data to")
    args = parser.parse_args()
    extract(args.input_file, args.output_file.name)
