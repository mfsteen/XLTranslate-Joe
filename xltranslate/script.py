#

import logging
import argparse

import openpyxl
import h5py

from . import wsp

log = logging.getLogger(__name__)

WSP_REGISTRY = {
    'Key Stats': wsp.key_stats.factory,
    'Income Statement': wsp.income_statement.factory,
    'Balance Sheet': wsp.balance_sheet.factory,
    'Cash Flow': wsp.cash_flow.factory,
    "Multiples": wsp.multiples.factory,
    "Historical Capitalization": wsp.historical_capitalization.factory,
    "Capital Structure Summary": wsp.capital_structure_summary.factory,
    "Ratios": wsp.ratios.factory,
    "Supplemental": wsp.supplemental.factory,
    "Industry Specific": wsp.industry_specific.factory,
    "Pension OPEB": wsp.pension_opeb.factory,
    "Segments": wsp.segments.factory,
}


def extract(input_file, output_file_name):
    wb = openpyxl.load_workbook(input_file, read_only=True)
    h5 = h5py.File(output_file_name, "w")
    for sheet in wb:
        factory = WSP_REGISTRY.get(sheet.title, None)
        if factory is not None:
            wspobj = factory(sheet)
            group = h5.create_group(sheet.title)
            wsp.to_hdf5(group, wspobj)
        else:
            log.info("Ignoring unknown sheet '%s'." % (sheet.title, ))


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
