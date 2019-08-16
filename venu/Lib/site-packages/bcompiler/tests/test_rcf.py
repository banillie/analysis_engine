from openpyxl import load_workbook
from ..analysers import rcf_run


def test_rcf(master_with_quarter_year_in_filename, tmpdir):
    wb_master = load_workbook(master_with_quarter_year_in_filename)
    wb_master.save(tmpdir.join(master_with_quarter_year_in_filename.split('/')[-1]))
    rcf_run(tmpdir, tmpdir)
    wb = load_workbook(tmpdir.join('PROJECT_PROGRAMME_NAME_1_RCF.xlsx'))
    ws = wb.active
    assert ws['B2'].value == "Reporting period (GMPP - Snapshot Date)"
    assert ws['C2'].value == "Approval MM1"


def test_rcf_chart_data_columns(master_with_quarter_year_in_filename, tmpdir):
    wb_master = load_workbook(master_with_quarter_year_in_filename)
    wb_master.save(tmpdir.join(master_with_quarter_year_in_filename.split('/')[-1]))
    rcf_run(tmpdir, tmpdir)
    wb = load_workbook(tmpdir.join('PROJECT_PROGRAMME_NAME_1_RCF.xlsx'))
    ws = wb.active
    assert ws['B10'].value == "SOBC"
    assert ws['C10'].value == 730
    assert ws['D10'].value == 1
    assert ws['B11'].value == "OBC"
    assert ws['C11'].value == 1826
    assert ws['D11'].value == 2
    assert ws['B12'].value == "FBC"
    assert ws['C12'].value is None  # need to verify why this passes
    assert ws['D12'].value == 3

