import datetime
import pytest

from openpyxl import load_workbook

from ..analysers.swimlane import run as swimlane_run
from ..analysers.swimlane_assurance import run as swimlane_assurance_run


def test_basic_swimlane_data(tmpdir, master):
    """
    This tests production of the spreadsheet containing the basic swimlane_milestones
    chart.

    This test tests the default implementation where value for day_range in config.ini
    is set to 365 - so a year from today's date.

    conftest.py generates a master which is used here. Some cells in the master are
    set to contain dates so they can be tested here.

    This test also relies on the following config.ini values:

        block_start = 90
        block_skip = 6
        block_end = 269
        forecast_actual_skip = 3
        milestones_to_collect = 30

    """
    today = datetime.date.today()
    tmpdir = [tmpdir]  # hacking the fact that output_path in implementation is list
    swimlane_run(output_path=tmpdir, user_provided_master_path=master)
    tmpdir = tmpdir[0]  # hacking the fact that output_path in implementation is list
    output = load_workbook(tmpdir.join('swimlane_milestones.xlsx'))
    ws = output.active
    assert ws['A1'].value == "PROJECT/PROGRAMME NAME 1"
    assert ws['A2'].value == "APPROVAL MM1 1" # config.ini: block_start row
    assert ws['A3'].value == "APPROVAL MM2 1"  # config.ini: block_start + block_skip
    assert ws['A4'].value == "APPROVAL MM3 1"  # config.ini: last one + block_skip
    assert ws['A5'].value == "APPROVAL MM4 1"  # config.ini: last one + block_skip
    assert ws['A6'].value == "APPROVAL MM5 1"  # config.ini: last one + block_skip
    assert ws['A7'].value == "APPROVAL MM6 1"  # config.ini: last one + block_skip
    assert ws['A8'].value == "APPROVAL MM7 1"  # config.ini: last one + block_skip
    assert ws['A9'].value == "APPROVAL MM8 1"  # config.ini: last one + block_skip
    assert ws['A10'].value == "APPROVAL MM9 1"  # config.ini: last one + block_skip

    assert ws['B2'].value == datetime.datetime(2015, 1, 1)
    assert ws['B3'].value == datetime.datetime(2019, 1, 1)
    assert ws['B4'].value == datetime.datetime(2020, 1, 1)
    assert ws['B5'].value == datetime.datetime(2020, 1, 1)
    assert ws['B6'].value == datetime.datetime(2020, 1, 1)
    assert ws['B7'].value == datetime.datetime(2020, 1, 1)
    assert ws['B8'].value == datetime.datetime(2020, 1, 1)
    assert ws['B9'].value == datetime.datetime(2020, 1, 1)
    assert ws['B10'].value == datetime.datetime(2020, 1, 1)

    assert ws['C2'].value is None
    assert ws['C3'].value is None
    assert ws['C4'].value is None
    assert ws['C5'].value is None
    assert ws['C6'].value is None
    assert ws['C7'].value is None
    assert ws['C8'].value is None
    assert ws['C9'].value is None
    assert ws['C10'].value is None


@pytest.mark.skip("VERIFY FAIL")
def test_swimlane_assurance_data(tmpdir, master):
    """
    This tests production of the spreadsheet containing the basic swimlane_milestones
    chart.

    This test tests the default implementation where value for day_range in config.ini
    is set to 365 - so a year from today's date.

    conftest.py generates a master which is used here. Some cells in the master are
    set to contain dates so they can be tested here.

    This test also relies on the following config.ini values:

        block_start = 1035
        block_skip = 6
        block_end = 1137
        forecast_actual_skip = 3
        milestones_to_collect = 30

    """
    today = datetime.date.today()
    tmpdir = [tmpdir]  # hacking the fact that output_path in implementation is list
    swimlane_assurance_run(output_path=tmpdir, user_provided_master_path=master)
    tmpdir = tmpdir[0]  # hacking the fact that output_path in implementation is list
    output = load_workbook(tmpdir.join('swimlane_assurance_milestones.xlsx'))
    ws = output.active
    assert ws['A1'].value == "PROJECT/PROGRAMME NAME 1"
    assert ws['A2'].value == "ASSURANCE MM1 1" # config.ini: block_start row
    assert ws['A3'].value == "ASSURANCE MM2 1"  # config.ini: block_start + block_skip
    assert ws['A4'].value == "ASSURANCE MM3 1"  # config.ini: last one + block_skip
    assert ws['A5'].value == "ASSURANCE MM4 1"  # config.ini: last one + block_skip
    assert ws['A6'].value == "ASSURANCE MM5 1"  # config.ini: last one + block_skip
    assert ws['A7'].value == "ASSURANCE MM6 1"  # config.ini: last one + block_skip
    assert ws['A8'].value == "ASSURANCE MM7 1"  # config.ini: last one + block_skip
    assert ws['A9'].value == "ASSURANCE MM8 1"  # config.ini: last one + block_skip
    assert ws['A10'].value == "ASSURANCE MM9 1"  # config.ini: last one + block_skip
    assert ws['A11'].value == "ASSURANCE MM10 1"  # config.ini: last one + block_skip

    assert ws['B2'].value == datetime.datetime(2015, 1, 1)
    assert ws['B3'].value == datetime.datetime(2019, 1, 1)
    assert ws['B4'].value == datetime.datetime(2020, 1, 1)
    assert ws['B5'].value == datetime.datetime(2020, 1, 1)
    assert ws['B6'].value == datetime.datetime(2020, 1, 1)
    assert ws['B7'].value == datetime.datetime(2020, 1, 1)
    assert ws['B8'].value == datetime.datetime(2020, 1, 1)
    assert ws['B9'].value == datetime.datetime(2020, 1, 1)
    assert ws['B10'].value == datetime.datetime(2020, 1, 1)

    assert ws['C2'].value is None
    assert ws['C3'].value is None
    assert ws['C4'].value is None
    assert ws['C5'].value is None
    assert ws['C6'].value is None
    assert ws['C7'].value is None
    assert ws['C8'].value is None
    assert ws['C9'].value is None
    assert ws['C10'].value is None
