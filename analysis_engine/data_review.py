# from datamaps.api import project_data_from_master
from collections import Counter, OrderedDict
from typing import Dict
from datamaps.process import Cleanser
from analysis_engine.data import (
    root_path,
    # get_project_info_data,
)
from openpyxl import load_workbook

def get_review_info_data(master_file: str) -> Dict:

    """ Taken from datamaps project_data_from_master api and adapted
    so dictionary nests keys viw row not column """

    wb = load_workbook(master_file)
    ws = wb.active
    for cell in ws["A"]:
        # we don't want to clean None...
        if cell.value is None:
            continue
        c = Cleanser(cell.value)
        cell.value = c.clean()
    review_dict = {}
    for row in ws.iter_rows(min_row=2):
        key_name = ""
        o = OrderedDict()
        for cell in row:
            if cell.column == 1:
                key_name = cell.value
                review_dict[key_name] = o
            else:
                val = ws.cell(row=1, column=cell.column).value
                review_dict[key_name][val] = cell.value
    # remove any "None" projects that were pulled from the master
    try:
        del review_dict[None]
    except KeyError:
        pass
    return review_dict


review_data = get_review_info_data(str(root_path) + "/data_review/DATA_REVIEW.xlsx")
