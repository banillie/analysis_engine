
from openpyxl import Workbook, load_workbook
from analysis.data import list_of_masters_all, root_path, baseline_bc_stamp
from analysis.engine_functions import get_quarter_stamp, baseline_information

def amend_baselines():

    bl_dict = {}

    baseline_info = load_workbook(root_path/'input/baseline_info_2.xlsx')
    for name in (list_of_masters_all[0].projects):
        '''worksheet is created for each project'''
        ws = baseline_info[name[0:29]]  # opening project ws

        #lower_dictionary = {}
        other_dict = {}

        for x in range(2, ws.max_column+1):
            lower_dictionary = {}
            quarter = ws.cell(row=1, column=x).value
            for i in range(2, ws.max_row):
                value = ws.cell(row=i, column=x).value
                key = ws.cell(row=i, column=1).value
                lower_dictionary[key] = value

            other_dict[quarter] = lower_dictionary

        bl_dict[name] = other_dict

    return bl_dict

run = amend_baselines()