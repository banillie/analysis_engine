"""
code in development which changes cost key names in masters.
"""

import typing
from typing import Dict

from openpyxl import load_workbook
from openpyxl.workbook import workbook

from data_mgmt.data import root_path, YEAR_LIST

from datamaps.api import project_data_from_master

key_change_master = load_workbook(root_path / "core_data/data_mgmt/key_change_log.xlsx")

test_id_master = "/home/will/code/python/analysis_engine/tests/resources/test_project_group_id_no.xlsx"

test_cost_masters_list = ["/home/will/code/python/analysis_engine/tests/resources/cost_test_master_1_2020.xlsx",
                          "/home/will/code/python/analysis_engine/tests/resources/cost_test_master_4_2019.xlsx",
                          "/home/will/code/python/analysis_engine/tests/resources/cost_test_master_4_2018.xlsx"]

test_masters_list = ["/home/will/code/python/analysis_engine/tests/resources/test_master_1_2020.xlsx",
                     "/home/will/code/python/analysis_engine/tests/resources/test_master_4_2019.xlsx",
                     "/home/will/code/python/analysis_engine/tests/resources/test_master_4_2018.xlsx",
                     "/home/will/code/python/analysis_engine/tests/resources/test_master_4_2017.xlsx"]

datamap_list = []

all_files = [test_masters_list, test_cost_masters_list]


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
    places milestone altered key names, in altered key names in dictionary, into master.
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
                ws.cell(row=row_num, column=1).value = year + " CDEL Forecast one off new costs"

    return wb


def run_change_keys() -> None:
    """
    runs code which replaces old key names with new names in master excel workbooks.
    """
    key_change_dict = put_keys_in_dict(key_change_master)
    for file_set in all_files:
        for f in file_set:
            changed_wb = alter_master_keys(f, key_change_dict)
            changed_wb.save(f)


def get_older_data(master_file: typing.TextIO, id_file: typing.TextIO) -> None:
    """
    Gets all old financial data across quarters and places into project id document.
    """
    master = project_data_from_master(master_file, 2, 2020)
    wb = load_workbook(id_file)
    ws = wb.active

    for i in range(1, ws.max_column + 1):
        project_name = ws.cell(row=1, column=1 + i).value
        for row_num in range(2, ws.max_row + 1):
            key = ws.cell(row=row_num, column=1).value
            try:
                if key in master.data[project_name].keys():
                    ws.cell(row=row_num, column=1 + i).value = master[project_name][key]
            except KeyError:  # project might not be present in quarter
                pass

    wb.save(id_file)


def run_get_old_data(id_master) -> None:
    for f in test_masters_list:
        get_older_data(f, id_master)


def place_old_data_in_master(master_file: typing.TextIO, id_file: typing.TextIO) -> None:
    """
    Gets all old financial data across quarters and places into project id document.
    """
    id_master = project_data_from_master(id_file, 2, 2020)
    wb = load_workbook(master_file)
    ws = wb.active

    for i in range(1, ws.max_column + 1):
        project_name = ws.cell(row=1, column=1 + i).value
        for row_num in range(2, ws.max_row + 1):
            key = ws.cell(row=row_num, column=1).value
            try:
                if key in id_master.data[project_name].keys():
                    ws.cell(row=row_num, column=1 + i).value = id_master[project_name][key]
            except KeyError:  # project might not be present in quarter
                pass

    wb.save(master_file)

def run_place_old_data_in_master(id_master) -> None:
    for f in test_cost_masters_list:
        place_old_data_in_master(f, id_master)


# run_get_old_data(test_id_master)
# run_change_keys()
run_place_old_data_in_master(test_id_master)
