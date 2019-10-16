'''
programme for altering across all masters changes in milestone keys.

working. bit of a hack at the mo though. further work required to finalise it and make pretty!

'''

#TODO get code to handle none data.

from openpyxl import load_workbook
from openpyxl.styles import Font
from analysis.data import list_of_masters_all
from analysis.engine_functions import ap_p_milestone_data_bulk, assurance_milestone_data_bulk, project_time_difference, \
    all_milestone_data_bulk, bc_ref_stages, master_baseline_index

def change_key(proj_list, q_master_wb_title_list, baseline_list, key_change_dict):
    red_text = Font(color="00fc2525")
    for proj_name in proj_list:
        for index in baseline_list[proj_name]:
            print(index)
            wb = load_workbook(q_master_wb_title_list[index])
            ws = wb.active
            for col_num in range(2, ws.max_column + 1):
                project_name = ws.cell(row=1, column=col_num).value
                #print(project_name)
                if project_name == proj_name:
                    print(proj_name)
                    for row_num in range(2, ws.max_row + 1):
                        for i in range(1, 6):
                            try:
                                if ws.cell(row=row_num, column=col_num).value == key_change_dict[proj_name]['Key '+ str(i)]:
                                    print(key_change_dict[proj_name]['Key '+ str(i)])
                                    ws.cell(row=row_num, column=col_num).value = key_change_dict[proj_name]['Key '+ str(i)+' change']
                                    print(key_change_dict[proj_name]['Key '+ str(i)+' change'])
                                    ws.cell(row=row_num, column=col_num).font = red_text
                                else:
                                    pass
                            except KeyError:
                                pass

            wb.save(q_master_wb_title_list[index])


'''INSTRUCTIONS FOR RUNNING THE PROGRAMME'''

'''1) load all master quarter data files here. They have to be store twice. The first time they are converted into
dictionaries. The second time the filepath is stored as a part of the list (this is so the current masters can be 
opened amended and saved again, make sure lists are identical'''


'''ii) list of file paths to masters'''
master_list = ('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2019_wip_(25_7_19).xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2018.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2018.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_2_2018.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2018.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2017.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2017.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_2_2017.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2017.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2016.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2016.xlsx')

'''2) Put the masters dictionaries into a list. '''

'''3) provide file path to document which contains information on the data that needs to be changed'''
key_change = project_data_from_master('C:\\Users\\Standalone\\general\\change_milestone_key_testing.xlsx')

'''4) list of projects. taken from the key change document - as this contains the only projects that need information 
changed'''
proj_list = list(key_change.keys())
proj_list_bespoke = ['South West Route Capacity', 'North of England Programme']

'''ignore this part. no change required'''
baseline_bc = bc_ref_stages(proj_list, list_of_masters_all)
q_master_baseline_no = master_baseline_index(proj_list, list_of_masters_all, baseline_bc)

'''5) enter relevant variables in the change_key function'''
change_key(proj_list, master_list, q_master_baseline_no, key_change)