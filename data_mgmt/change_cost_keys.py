"""
code in development which changes cost key names in masters.
"""


import typing
from typing import Dict

from openpyxl import load_workbook

from openpyxl.workbook import workbook

from data_mgmt.data import root_path

from data_mgmt.data import YEAR_LIST

key_change_master = load_workbook(root_path / "core_data/data_mgmt/key_change_log.xlsx")

masters_list = ["/home/will/code/python/analysis_engine/tests/resources/cost_test_master_1_2020.xlsx",
                "/home/will/code/python/analysis_engine/tests/resources/cost_test_master_4_2019.xlsx",
                "/home/will/code/python/analysis_engine/tests/resources/cost_test_master_4_2018.xlsx"]


def put_keys_in_dict(wb: workbook) -> Dict[str, str]:
    """
    places key information i.e. keys old and new names from wb into a python dictionary
    """
    ws = wb.active
    output_dict = {}

    for x in range(1, ws.max_row + 1):
        key = ws.cell(row=x, column=1).value
        codename = ws.cell(row=x, column=2).value
        output_dict[key] = codename

    return output_dict


def alter_master_keys(file: typing.TextIO, keys_dict: Dict[str, str]) -> workbook:
    """
    places milestone altered key names, in altered key names dictionary, into master.
    """
    wb = load_workbook(file)
    ws = wb.active

    for row_num in range(2, ws.max_row + 1):
        for key in keys_dict.keys():  # changes stored in the altered key change log wb
            if ws.cell(row=row_num, column=1).value == key:
                ws.cell(row=row_num, column=1).value = \
                    keys_dict[key]
        for year in YEAR_LIST:  # changes to yearly profile keys
            if ws.cell(row=row_num, column=1).value == year + " CDEL Forecast Total":
                print('match')
                ws.cell(row=row_num, column=1).value = year + " CDEL Forecast one off new costs"

    return wb


key_change_dict = put_keys_in_dict(key_change_master)

for f in masters_list:
    changed_wb = alter_master_keys(f, key_change_dict)
    changed_wb.save(f)


