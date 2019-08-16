import os
import tempfile
from datetime import date

from openpyxl import load_workbook

import bcompiler.compile as compile_module
from bcompiler.utils import runtime_config as config
from ..compile import parse_comparison_master
from ..compile import parse_source_cells as parse
from ..compile import run
from ..process.datamap import Datamap

TODAY = date.today().isoformat()
TEMPDIR = tempfile.gettempdir()

AUX_DIR = "/".join([TEMPDIR, 'bcompiler'])
SOURCE_DIR = "/".join([AUX_DIR, 'source'])
RETURNS_DIR = "/".join([SOURCE_DIR, 'returns'])
OUTPUT_DIR = "/".join([AUX_DIR, 'output'])

setattr(compile_module, 'RETURNS_DIR', RETURNS_DIR)
setattr(compile_module, 'OUTPUT_DIR', OUTPUT_DIR)
setattr(compile_module, 'TODAY', date.today().isoformat())

q_string = config['QuarterData']['CurrentQuarter'].split()[0]


def test_populate_single_template_from_master(populated_template, datamap):
    """
    This tests bcompiler -b X essentially (or bcompiler -a, but for a single file.
    """
    data = parse(populated_template, datamap)
    assert data[0]['gmpp_key'] == 'Project/Programme Name'
    assert data[0]['gmpp_key_value'] == 'PROJECT/PROGRAMME NAME 9'


def test_compile_all_returns_to_master_no_comparison(populated_template, datamap):
    """
    This tests 'bcompiler compile' or 'bcompiler' option.
    """
    # print([item for item in dir(compile_module) if not item.startswith("__")])
    # patching module attributes to get it working
    setattr(compile_module, 'DATAMAP_RETURN_TO_MASTER', datamap)
    run()
    # for one of the templates that we have compiled (using 9 I think)...
    # get the project title...
    data = parse(populated_template, datamap)
    project_title = data[0]['gmpp_key_value']
    # then we need to open up the master that was produced by run() function above...
    wb = load_workbook(os.path.join(OUTPUT_DIR, 'compiled_master_{}_{}.xlsx'.format(TODAY, q_string)))
    ws = wb.active
    # we then need the "Project/Programme Name" row from the master
    project_title_row = [i.value for i in ws[1]]
    assert project_title in project_title_row


def test_compile_all_returns_to_master_with_string_comparison(datamap, previous_quarter_master, populated_template_comparison):
    """
    This tests 'bcompiler compile --compare' or 'bcompiler --compare' option.
    :param populated_template:
    :param datamap:
    :param previous_quarter_master:
    :return:
    """
    setattr(compile_module, 'DATAMAP_RETURN_TO_MASTER', datamap)
    comparitor = parse_comparison_master(previous_quarter_master)
    run(comparitor=comparitor)
    # now to test the cell styling to make sure it's changed
    wb = load_workbook(os.path.join(OUTPUT_DIR, 'compiled_master_{}_{}.xlsx'.format(TODAY, q_string)))
    ws = wb.active
    # We need to gather the cells from row 11, and compare WORKING CONTACT NAME 1, 2 and 3
    working_contact_row = [i for i in ws[11]]
    # checking for yellow background characteristic of a changed string
    assert working_contact_row[1].value == "WORKING CONTACT NAME 2"
    assert working_contact_row[1].fill.bgColor.rgb == '00FCF5AA'
    assert working_contact_row[2].value == "WORKING CONTACT NAME 1"
    assert working_contact_row[2].fill.bgColor.rgb == '00FCF5AA'
    assert working_contact_row[3].value == "WORKING CONTACT NAME 0"
    # testing default 000000 background
    assert working_contact_row[0].fill.bgColor.rgb == '00000000'


def test_compile_all_returns_to_master_with_date_comparison(datamap, previous_quarter_master, populated_template_comparison):
    """
    This depends upon the fixture setting an earlier date in the previous_quarter_master.
    :param datamap:
    :param previous_quarter_master:
    :return:
    """
    setattr(compile_module, 'DATAMAP_RETURN_TO_MASTER', datamap)
    comparitor = parse_comparison_master(previous_quarter_master)
    run(comparitor=comparitor)
    # now to test the cell styling to make sure it's changed
    wb = load_workbook(os.path.join(OUTPUT_DIR, 'compiled_master_{}_{}.xlsx'.format(TODAY, q_string)))
    ws = wb.active
    # we need to find reference for "SRO Tenure Start Date"

    # we know it's row 13, but what column? index of where "PROJECT/PROGRAMME NAME 1" in row 1
    project_title_row = [i.value for i in ws[1]]

    # testing for a earlier (green) colour now
    target_index = [project_title_row.index(i) for i in project_title_row if i == 'PROJECT/PROGRAMME NAME 1'][0]
    target_cell = ws.cell(row=85, column=target_index + 1) # take into account zero indexing

    # comparison code is at cellformat.py:135
    assert target_cell.fill.bgColor.rgb == '00ABFCA9' # LIGHT GREEN because THIS value is HIGHER/LATER than comp

#   # testing for a later (violet) colour now
#   target_index = [project_title_row.index(i) for i in project_title_row if i == 'PROJECT/PROGRAMME NAME 2'][0]
#   target_cell = ws.cell(row=14, column=target_index + 1) # take into account zero indexing
#
#   # comparison code is at cellformat.py:241
#   assert target_cell.fill.bgColor.rgb == '00A9AAFC' # LIGHT GREEN because THIS value is HIGHER/LATER than comp


def test_datamap_class(datamap):
    """
    This tests correct creation of Datamap object.
    """
    dm = Datamap()
    dm.cell_map_from_csv(datamap)
    assert dm.cell_map[1].cell_reference == 'B49'
