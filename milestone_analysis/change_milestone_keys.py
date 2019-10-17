'''
programme for altering across all masters changes in milestone keys.

Issue: not currently marking new text as red. not clear why this line of code not working

code doesn't/can't handle keys that have been removed.
'''


from datamaps.api import project_data_from_master
from openpyxl import load_workbook
from openpyxl.styles import Font
from analysis.data import list_of_masters_all, q2_1920
from analysis.engine_functions import ap_p_milestone_data_bulk, assurance_milestone_data_bulk, project_time_difference, \
    all_milestone_data_bulk, bc_ref_stages, master_baseline_index

def change_key(project_list, master_wb_title_list, baseline_list, change_key):
    '''
    :param project_list: list of project names
    :param master_wb_title_list: list of quarter master data workbooks. This is different from the master list of
    python dictionaries usually passed into functions.
    :param baseline_list: list which indexes where baseline information sits.
    :param change_key: list of keys that need to be changed.
    :return: the wbs in the master_wb_title_list with the data amended.
    '''

    red_text = Font(color="00fc2525")

    for name in project_list:
        print(name)
        for index in baseline_list[name]:
            print(index)
            wb = load_workbook(master_wb_title_list[index])
            ws = wb.active
            for col_num in range(2, ws.max_column + 1):
                project_name = ws.cell(row=1, column=col_num).value
                if project_name == name:
                    for row_num in range(2, ws.max_row + 1):
                        for i in range(1, 11):
                            try:
                                if ws.cell(row=row_num, column=col_num).value == change_key[project_name]['Key '+ str(i)]:
                                    print(change_key[project_name]['Key '+ str(i)])
                                    ws.cell(row=row_num, column=col_num).value = change_key[project_name]['Key '+ str(i)+' change']
                                    print(change_key[project_name]['Key '+ str(i)+' change'])
                                    ws.cell(row=row_num, column=col_num).font = red_text
                                else:
                                    pass
                            except KeyError:
                                pass

            wb.save(master_wb_title_list[index])


'''INSTRUCTIONS FOR RUNNING THE PROGRAMME'''

'''ONE. List of file paths to masters'''
master_list = ('C:\\Users\\Standalone\\general\\core_data\\master_2_2019.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_1_2019.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_4_2018.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_3_2018.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_2_2018.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_1_2018.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_4_2017.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_3_2017.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_2_2017.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_1_2017.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_4_2016.xlsx',
               'C:\\Users\\Standalone\\general\\core_data\\master_3_2016.xlsx')

'''TWO. Provide file path to document which contains information on the data that needs to be changed'''
key_change = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\project_milestones\\'
                                      'master_key_change_record_q2_1920.xlsx', 2, 2019)

'''THREE. List of projects. taken from the key change document - as this contains the only projects that need 
information changed'''
project_name_list = key_change.projects

'''ignore this part. no change required'''
baseline_bc = bc_ref_stages(project_name_list, list_of_masters_all)
master_baseline_no = master_baseline_index(project_name_list, list_of_masters_all, baseline_bc)

'''FOUR. enter relevant variables in the change_key function'''
change_key(project_name_list, master_list, master_baseline_no, key_change)