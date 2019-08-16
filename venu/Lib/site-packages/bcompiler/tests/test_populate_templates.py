import configparser
import os
import tempfile
from datetime import datetime

from openpyxl import load_workbook

import bcompiler.main as main_module
from ..core import Quarter, Master
from ..main import get_list_projects
from ..main import populate_blank_bicc_form as populate
from ..utils import project_data_from_master

TEMPDIR = tempfile.gettempdir()

AUX_DIR = "/".join([TEMPDIR, 'bcompiler'])
SOURCE_DIR = "/".join([AUX_DIR, 'source'])
RETURNS_DIR = "/".join([SOURCE_DIR, 'returns'])
OUTPUT_DIR = "/".join([AUX_DIR, 'output'])

config = configparser.ConfigParser()
CONFIG_FILE = 'test_config.ini'
config.read(CONFIG_FILE)
current_quarter = config['QuarterData']['CurrentQuarter']


def test_get_list_projects_main_xlsx(master):
    l = get_list_projects(master)
    assert l[0] == 'PROJECT/PROGRAMME NAME 1'


def test_pull_data_from_xlsx_master(master):
    data = project_data_from_master(master)
    assert data['PROJECT/PROGRAMME NAME 1']['SRO Sign-Off'] == 'SRO SIGN-OFF 1'
    assert data['PROJECT/PROGRAMME NAME 1'][
        'Reporting period (GMPP - Snapshot Date)'] == 'REPORTING PERIOD (GMPP - SNAPSHOT DATE) 1'


def test_populate_single_template(master, blank_template, datamap):
    setattr(main_module, 'OUTPUT_DIR', OUTPUT_DIR)
    setattr(main_module, 'SOURCE_DIR', SOURCE_DIR)
    setattr(main_module, 'BLANK_TEMPLATE_FN', ''.join(['/', blank_template.split('/')[-1]]))
    m = Master(Quarter(3, 2017), master)
    populate(m, 0)
    wb = load_workbook(os.path.join(OUTPUT_DIR, f'PROJECT_PROGRAMME NAME 1_{current_quarter}_Return.xlsm'))
    ws = wb[config['TemplateTestData']['summary_sheet']]
    assert ws['B5'].value == 'PROJECT/PROGRAMME NAME 1'
    # for f in glob.glob('/'.join([OUTPUT_DIR, '*_Return.xlsm'])):
    #     os.remove(f)


def test_populate_date_cell(master, blank_template):
    setattr(main_module, 'OUTPUT_DIR', OUTPUT_DIR)
    setattr(main_module, 'SOURCE_DIR', SOURCE_DIR)
    setattr(main_module, 'BLANK_TEMPLATE_FN', ''.join(['/', blank_template.split('/')[-1]]))
    m = Master(Quarter(3, 2017), master)
    populate(m, 0)
    wb = load_workbook(os.path.join(OUTPUT_DIR, f'PROJECT_PROGRAMME NAME 1_{current_quarter}_Return.xlsm'))
    ws = wb[config['TemplateTestData']['fb_sheet']]
    assert ws['E11'].value == datetime(2017, 6, 20)
    assert ws['C13'].value == datetime(2017, 6, 20)
    ws = wb[config['TemplateTestData']['summary_sheet']]
    assert ws['C15'].value == datetime(2017, 8, 10)
    # for f in glob.glob('/'.join([OUTPUT_DIR, '*_Return.xlsm'])):
    #     os.remove(f)
