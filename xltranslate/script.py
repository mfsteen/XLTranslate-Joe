#

import logging
import argparse

import openpyxl

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


def extract(input_file):
    wb = openpyxl.load_workbook(input_file, read_only=True)
    for sheet in wb:
        klass = WSP_REGISTRY.get(sheet.title, None)
        if klass is not None:
            wspobj = klass(sheet)
            print("\n")
            print("="*79)
            print("%s\n" % (sheet.title, ))
            wspobj.dump()
        else:
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
