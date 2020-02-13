'''returns data required by no 10 commission'''

from openpyxl import Workbook
from openpyxl.styles import PatternFill
from analysis.data import list_of_masters_all, latest_quarter_project_names, bc_index, baseline_bc_stamp, salmon_fill, \
    root_path
from analysis.engine_functions import all_milestone_data_bulk, ap_p_milestone_data_bulk, assurance_milestone_data_bulk,\
    get_all_project_names, get_quarter_stamp, grey_conditional_formatting

def return_baseline_data(project_name_list, data_key_list):

    wb = Workbook()

    for i, project_name in enumerate(project_name_list):
        '''worksheet is created for each project'''
        ws = wb.create_sheet(project_name, i)  # creating worksheets
        ws.title = project_name  # title of worksheet

        '''list project names, groups and stage in ws'''
        for y, key in enumerate(data_key_list):
            ws.cell(row=2+y, column=1, value=key)
            for x in range(0, len(baseline_bc_stamp[project_name])):
                index = baseline_bc_stamp[project_name][x][2]
                print(index)
                try:
                    ws.cell(row=2+y, column=2+x, value=list_of_masters_all[index].data[project_name][key])
                except KeyError:
                    ws.cell(row=2+y, column=2+x, value='not reporting')

    return wb

def return_data(project_name_list, data_key_list):

    wb = Workbook()

    for i, project_name in enumerate(project_name_list):
        '''worksheet is created for each project'''
        ws = wb.create_sheet(project_name, i)  # creating worksheets
        ws.title = project_name  # title of worksheet

        '''list project names, groups and stage in ws'''
        for y, key in enumerate(data_key_list):
            ws.cell(row=2+y, column=1, value=key)
            for x, master in enumerate(list_of_masters_all):
                try:
                    ws.cell(row=2+y, column=2+x, value=master.data[project_name][key])
                except KeyError:
                    ws.cell(row=2+y, column=2+x, value='not reporting')

    return wb


milestone_data_interest = ['BICC approval point', 'Total Forecast']

'''THREE. Run the programme'''
'''option one - run the return_milestone_data for all milestone data'''
run_standard = return_baseline_data(latest_quarter_project_names, milestone_data_interest)

'''FOUR. specify the file path and name of the output document'''
run_standard.save(root_path/'output/no_10_data.xlsx')
