import csv
import os
import subprocess

from bcompiler import __version__
from ..utils import OUTPUT_DIR


def test_bcompiler_help():
    output = subprocess.run(['bcompiler', '-h'], stdout=subprocess.PIPE, encoding='utf-8')
    assert output.stdout.startswith('usage')


def test_bcompiler_version():
    output = subprocess.run(['bcompiler', '-v'], stdout=subprocess.PIPE, encoding='utf-8')
    assert output.stdout.strip() == __version__


def test_bcompiler_count_rows(populated_template):
    output = subprocess.run(['bcompiler', '-r'], stdout=subprocess.PIPE, encoding='utf-8')
    assert output.stdout.startswith('Workbook')


def test_bcompiler_count_rows_csv(populated_template):
    subprocess.run(['bcompiler', '-r', '--csv'])
    with open(os.path.join(OUTPUT_DIR, 'row_count.csv'), 'r') as f:
        reader = csv.reader(f)
        assert next(reader)[0] == 'bicc_template.xlsm'


def test_bcompiler_count_rows_quiet(populated_template):
    output = subprocess.run(['bcompiler', '-r', '--quiet'], stdout=subprocess.PIPE, encoding='utf-8')
    assert output.stdout.startswith('#')


#def test_bcompiler_populate_all_templates(master):
#    output = subprocess.run(['bcompiler', '-a'], stdout=subprocess.PIPE, encoding='utf-8')
#    assert output.stdout
