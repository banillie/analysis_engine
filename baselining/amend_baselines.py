
from openpyxl import load_workbook
from analysis.data import list_of_masters_all, root_path, red_text

def get_data(wb):
    '''Function that takes data from a wb across all ws in use and places them into
     a python dictionary
     wb: workbook containing the data
     returns and python dictionary'''

    bl_dict = {}

    for name in (list_of_masters_all[0].projects):
        '''worksheet is created for each project'''
        ws = wb[name[0:29]]  # opening project ws

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

def place_in_masters(project_list, master_wb_title_list, baseline_data):
    '''
    Places data into masters and removes unnecessary data by putting in blanks.
    :param project_list: list of project names
    :param master_wb_title_list: list of quarter master data workbooks. This is different from the master list of
    python dictionaries usually passed into functions.
    :param baseline_data: list of keys that need to be changed.
    :return: the wbs in the master_wb_title_list with the data amended.
    '''

    for name in project_list:
        for master in master_wb_title_list:
            wb = load_workbook(master)
            ws = wb.active
            for col_num in range(2, ws.max_column + 1):
                project_name = ws.cell(row=1, column=col_num).value
                if project_name == name:
                    print(name)
                    project_bl_dict = baseline_data[name]
                    for row_num in range(2, ws.max_row + 1):
                        if ws.cell(row=row_num, column=1).value == 'Reporting period (GMPP - Snapshot Date)':
                            quarter = ws.cell(row=row_num, column=col_num).value
                            project_bl_dict_keys = list(project_bl_dict[quarter].keys())
                    # some hard coding required as key names between master and bl_data not a match
                    # to sort for next iteration
                    for i in range(len(project_bl_dict_keys)):
                        for row_num in range(2, ws.max_row + 1):
                            wb_key = ws.cell(row=row_num, column=1).value
                            if wb_key == 'IPDC approval point':
                                # ws.cell(row=row_num, column=col_num).value = project_bl_dict[quarter]['IPDC BC approval']
                                # ws.cell(row=row_num, column=col_num).font = red_text
                                print(project_bl_dict[quarter]['IPDC BC approval'])


                                    # baseline_data[quarter]:
                                    # print(change_key[project_name]['Key '+ str(i)])
                                    # ws.cell(row=row_num, column=col_num).value = change_key.data[project_name]['Key '+ str(i)+' change']
                                    # print(change_key[project_name]['Key '+ str(i)+' change'])
                                    # ws.cell(row=row_num, column=col_num).font = red_text
                else:
                    pass

            wb.save(master)

master_list = [root_path/'core_data/master_4_2019.xlsx']

baseline_data_wb = load_workbook(root_path / 'input/baseline_info_2.xlsx')

bl_data = get_data(baseline_data_wb)

place_in_masters(list_of_masters_all[0].projects, master_list, bl_data)