"""
Test some functions of the script.
"""
from openpyxl import load_workbook
from xlsx_to_json import processsheet

TEST_FILE = "sample/ResourceAssessmentSummaryData032011.xlsx"
EXPECTED_SHEET_NAMES = ['All Summary Data']

def test_processsheet():
    """
    Compare that the output has 191 rows of data with 15 cols.
    """
    workbook = load_workbook(TEST_FILE, read_only=True)
    sheet = EXPECTED_SHEET_NAMES[0]
    data = processsheet(workbook, sheet)
    assert len(data) == 191
    sample = data[0]
    assert len(sample) == 15
