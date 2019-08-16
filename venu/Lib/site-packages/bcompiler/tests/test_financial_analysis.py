import os

import pytest
from openpyxl import load_workbook

from ..analysers.financial import run as financial_run


@pytest.fixture
def master_repository(master, previous_quarter_master, tmpdir):
    with open(master, 'w') as f:
        f.write(os.path.join(tmpdir.dirname, master))
    with open(previous_quarter_master, 'w') as f:
        f.write(os.path.join(tmpdir.dirname, previous_quarter_master))
    return tmpdir.strpath


def test_a1_cell(tmpdir, master_repository):
    tmpdir = [tmpdir]
    financial_run(master_repository, output_path=tmpdir)
    tmpdir = tmpdir[0]
    wb = load_workbook(tmpdir.join('financial_analysis.xlsx'))
    ws = wb.get_sheet_by_name('PROJECT_PROGRAMME NAME 1')
    assert ws['A1'].value == 'PROJECT/PROGRAMME NAME 1'


def test_header_cells(tmpdir, master_repository):
    financial_run(master_repository, output_path=tmpdir)
    tmpdir = tmpdir[0]
    wb = load_workbook(tmpdir.join('financial_analysis.xlsx'))
    ws = wb.get_sheet_by_name('PROJECT_PROGRAMME NAME 1')
    assert ws['A3'].value == 'Q1 18/19'
