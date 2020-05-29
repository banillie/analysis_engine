'''

This programme calculates the time difference between reported milestones

Output document:
1) one excel workbook contain all project milestone information.

See instructions below.

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.

'''

from openpyxl import Workbook
from analysis.engine_functions import project_time_difference, all_milestones_dict
from analysis.data import list_of_masters_all, root_path, ipdc_milestone_bl_index, hsmrpg

def put_into_wb_all(project_name_list, t_data, td_data, td_data_two, wb):
    '''

    Function that places all data into excel wb for this programme

    project_name_list: list of project to return data for
    t_data: dictionary containing milestone data for projects.
    dictionary structure is {'project name': {'milestone name': datetime.date: 'notes'}}
    td_data: dictionary containing time_delta milestone data for projects.
    dictionary structure is {'project name': {'milestone name': 'time delta info'}}
    td_data_two: dictionary containing second time_delta data for projects.
    same structure as for td_data.
    wb: blank excel wb

    '''

    ws = wb.active

    row_num = 2
    for project_name in project_name_list:
        for i, milestone in enumerate(td_data[project_name].keys()):
            ws.cell(row=row_num + i, column=1).value = project_name
            ws.cell(row=row_num + i, column=2).value = milestone
            try:
                milestone_date = tuple(t_data[project_name][milestone])[0]
                ws.cell(row=row_num + i, column=3).value = milestone_date
                ws.cell(row=row_num + i, column=3).number_format = 'dd/mm/yy'
            except KeyError:
                ws.cell(row=row_num + i, column=3).value = 0

            try:
                value = td_data[project_name][milestone]
                ws.cell(row=row_num + i, column=4).value = value
            except KeyError:
                ws.cell(row=row_num + i, column=4).value = 0

            try:
                value = td_data_two[project_name][milestone]
                ws.cell(row=row_num + i, column=5).value = value
            except KeyError:
                ws.cell(row=row_num + i, column=5).value = 0

            try:
                milestone_date = tuple(t_data[project_name][milestone])[0]
                ws.cell(row=row_num + i, column=6).value = t_data[project_name][milestone][milestone_date]  # provides notes
            except IndexError:
                ws.cell(row=row_num + i, column=6).value = 0

        row_num = row_num + len(td_data[project_name])

    ws.cell(row=1, column=1).value = 'Project'
    ws.cell(row=1, column=2).value = 'Milestone'
    ws.cell(row=1, column=3).value = 'Date'
    ws.cell(row=1, column=4).value = '3/m change (days)'
    ws.cell(row=1, column=5).value = 'Baseline change (days)'
    ws.cell(row=1, column=6).value = 'Notes'

    return wb

def run_milestone_comparator(function, project_name_list, masters_list):

    '''
    Function that runs this programme.

    function: The type of milestone you wish to analysis can be specified through choosing all_milestone_data_bulk,
    ap_p_milestone_data_bulk, or assurance_milestone_data_bulk functions, all available from engine_function import
    statement above.
    project_name_list: list of project to return data for
    masters_list: list of masters containing quarter information
    date_of_interest: the date after which project milestones should be returned.

    '''


    wb = Workbook()

    '''gather mini-dictionaries for each quarter'''

    current_milestones_data = {}
    last_milestones_data = {}
    oldest_milestones_data = {}
    for project_name in project_name_list:
        #print(project_name)
        p_current_milestones_data = function([project_name], masters_list[0])
        current_milestones_data.update(p_current_milestones_data)
        p_last_milestones_data = function([project_name], masters_list[1])
        last_milestones_data.update(p_last_milestones_data)
        p_oldest_milestones_data = function([project_name], masters_list[ipdc_milestone_bl_index[project_name][2]])
        oldest_milestones_data.update(p_oldest_milestones_data)

    '''calculate time current and last quarter'''
    first_diff_data = project_time_difference(current_milestones_data, last_milestones_data)
    second_diff_data = project_time_difference(current_milestones_data, oldest_milestones_data)

    run = put_into_wb_all(project_name_list, current_milestones_data, first_diff_data, second_diff_data, wb)

    return run

''' RUNNING PROGRAMME '''

'''choose the type of variables that you would like to place in run_milestone_comparator function. Arguments 
are placed in this order. 

1. function: The type of milestone you wish to analysis can be specified through choosing all_milestone_data_bulk, 
ap_p_milestone_data_bulk, or assurance_milestone_data_bulk functions, all available from engine_function import 
statement above. 
2. project_name_list: list of project to return data for
3. masters_list: list of masters containing quarter information
 
'''
run = run_milestone_comparator(all_milestones_dict, hsmrpg, list_of_masters_all)

'''specify file path to output document'''
run.save(root_path/'output/portfolio_milestone_analysis_test.xlsx')