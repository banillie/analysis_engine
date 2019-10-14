'''

This programme calculates the time difference between reported milestones

Output document:
1) excel workbook with all project milestone information for each project

See instructions below.

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.

'''

#TODO solve problem re filtering in excle when values have + sign in front of the them

import datetime
from openpyxl import Workbook
from analysis.engine_functions import all_milestone_data_bulk, ap_p_milestone_data_bulk, assurance_milestone_data_bulk, \
    project_time_difference, bc_ref_stages, master_baseline_index, filter_project_group
from analysis.data import q2_1920, list_of_masters_all


def put_into_wb_all_single(project_name, t_data, td_data, td_data_two, baseline_record):
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

    wb = Workbook()
    ws = wb.active

    row_num = 2
    for i, milestone in enumerate(td_data[project_name].keys()):
        ws.cell(row=row_num + i, column=1).value = project_name
        ws.cell(row=row_num + i, column=2).value = milestone
        try:
            milestone_date = tuple(t_data[project_name][milestone])[0]
            ws.cell(row=row_num + i, column=3).value = milestone_date
        except KeyError:
            ws.cell(row=row_num + i, column=3).value = 0

        try:
            value = td_data[project_name][milestone]
            try:
                if int(value) > 0:
                    ws.cell(row=row_num + i, column=4).value = '+' + str(value) + ' (days)'
                elif int(value) < 0:
                    ws.cell(row=row_num + i, column=4).value = str(value) + ' (days)'
                elif int(value) == 0:
                    ws.cell(row=row_num + i, column=4).value = value
            except ValueError:
                ws.cell(row=row_num + i, column=4).value = value
        except KeyError:
            ws.cell(row=row_num + i, column=4).value = 0

        try:
            value = td_data_two[project_name][milestone]
            try:
                if int(value) > 0:
                    ws.cell(row=row_num + i, column=5).value = '+' + str(value) + ' (days)'
                elif int(value) < 0:
                    ws.cell(row=row_num + i, column=5).value = str(value) + ' (days)'
                elif int(value) == 0:
                    ws.cell(row=row_num + i, column=5).value = value
            except ValueError:
                ws.cell(row=row_num + i, column=5).value = value
        except KeyError:
            ws.cell(row=row_num + i, column=5).value = 0

        try:
            milestone_date = tuple(t_data[project_name][milestone])[0]
            ws.cell(row=row_num + i, column=6).value = t_data[project_name][milestone][milestone_date] # provided notes
        except IndexError:
            ws.cell(row=row_num + i, column=6).value = 0

    ws.cell(row=1, column=1).value = 'Project'
    ws.cell(row=1, column=2).value = 'Milestone'
    ws.cell(row=1, column=3).value = 'Date'
    ws.cell(row=1, column=4).value = '3/m change (days)'
    ws.cell(row=1, column=5).value = 'baseline change (days)'
    ws.cell(row=1, column=6).value = 'Notes'

    ws.cell(row=1, column=8).value = 'data baseline quarter'
    ws.cell(row=2, column=8).value = baseline_record[project_name][2][0]

    return wb

def run_milestone_comparator_single(function, project_name, masters_list):
    '''
    Function that runs this programme.

    function: The type of milestone you wish to analysis can be specified through choosing all_milestone_data_bulk,
    ap_p_milestone_data_bulk, or assurance_milestone_data_bulk functions, all available from engine_function import
    statement above.
    project_name_list: list of project to return data for
    masters_list: list of masters containing quarter information
    date_of_interest: the date after which project milestones should be returned.

    '''

    '''firstly business cases of interest are filtered out by bc_ref_stage function'''
    baseline_bc = bc_ref_stages([project_name], masters_list)
    baseline_list_index = master_baseline_index([project_name], masters_list, baseline_bc)

    '''project milestone data is captured across different quarters'''
    p_current_milestones = function([project_name], masters_list[baseline_list_index[project_name][0]])
    p_last_milestones = function([project_name], masters_list[baseline_list_index[project_name][1]])
    p_oldest_milestones = function([project_name], masters_list[baseline_list_index[project_name][2]])

    '''calculate time current and last quarter'''
    first_diff_data = project_time_difference(p_current_milestones, p_last_milestones)
    second_diff_data = project_time_difference(p_current_milestones, p_oldest_milestones)

    run = put_into_wb_all_single(project_name, p_current_milestones, first_diff_data, second_diff_data, baseline_bc)

    return run


''' RUNNING THE PROGRAMME'''

'''Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in 
the import statement.'''

''' ONE. set list of projects to be included in output'''
'''option one - all projects'''
project_quarter_list = q2_1920.projects

'''option two - group of projects. use filter_project_group function'''
project_group_list = filter_project_group(q2_1920, 'HSMRPG')

'''option three - single project'''
one_project_list = ['High Speed Rail Programme (HS2)']

'''TWO. the following for statement prompts the programme to run. 
step one - place the list of projects chosen in step three at the end of the for statement. i.e. for project_name in [here] 
step two - chose the variables required for the run_milestone_comparator_single function. The first argument in relation
to which milestone data is to be analysed will normally be the only change. 
step three - provide relevant file path to document output. Changing the quarter stamp info as necessary. Note keep {} in 
file name as this is where the project name is recorded in the file title'''
for project_name in project_group_list:
    print('Doing milestone movement analysis for ' + str(project_name))
    wb = run_milestone_comparator_single(all_milestone_data_bulk, project_name, list_of_masters_all)
    wb.save('C:\\Users\\Standalone\\general\\masters folder\\project_milestones\\'
            'q2_1920_{}_milestone_movement_analysis.xlsx'.format(project_name))