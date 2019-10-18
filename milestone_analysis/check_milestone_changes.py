'''

This programme checks if there has been changes in reported milestone keys.

Output document:
1) excel workbook with all project milestone information for each project

See instructions below.

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.

'''

import datetime
from openpyxl import Workbook
from openpyxl.styles import Font
from analysis.engine_functions import all_milestone_data_bulk, ap_p_milestone_data_bulk, assurance_milestone_data_bulk, \
    bc_ref_stages, master_baseline_index, filter_project_group
from analysis.data import q2_1920, list_of_masters_all

def check_m_keys_in_excel_single(project_name, t_data_one, t_data_two, t_data_three):
    '''
    function for placing all information into an excel wb

    :param project_name: project name
    :param t_data_one: dictionary containing latest milestone data for projects.
    dictionary structure is {'project name': {'milestone name': datetime.date: 'notes'}}
    :param t_data_two: dictionary containing last milestone date for projects. same structure as above.
    :param t_data_three: dictionary containing baseline milestone date for projects. same structure as above.

    '''

    wb = Workbook()
    ws = wb.active
    red_text = Font(color="00fc2525")

    row_num = 2

    one = list(t_data_one[project_name].keys())
    [x for x in one if x is not None].sort()
    two = list(t_data_two[project_name].keys())
    [x for x in two if x is not None].sort()
    three = list(t_data_three[project_name].keys())
    [x for x in three if x is not None].sort()

    long = longest_list(one, two, three)
    for i in range(0, len(long)):
        ws.cell(row=row_num + i, column=1).value = project_name
        try:
            ws.cell(row=row_num + i, column=2).value = one[i]
        except IndexError:
            pass
        try:
            ws.cell(row=row_num + i, column=3).value = two[i]
            if two[i] not in one:
                ws.cell(row=row_num + i, column=3).font = red_text
        except IndexError:
            pass
        try:
            ws.cell(row=row_num + i, column=4).value = three[i]
            if three[i] not in one:
                ws.cell(row=row_num + i, column=4).font = red_text
        except IndexError:
            pass


    row_heading_list = ['Project', 'This quarter', 'Last quarter', 'Baseline quarter', 'KEY MATCH']
    for i, heading in enumerate(row_heading_list):
        ws.cell(row=1, column=i+1).value = heading

    column_ltr_list = ['A', 'B', 'C', 'D', 'E']
    for ltr in (column_ltr_list):
        ws.column_dimensions[ltr].width = 40

    return wb

def longest_list(one, two, three):
    '''
    Function. Helper for the above. Calculates the longest list and therefore the one to use for iteration.

    :param one: list_one
    :param two: list_two
    :param three: list_three

    Returns the longest list.
    '''

    list_list = [one, two, three]
    a = len(one)
    b = len(two)
    c = len(three)

    out = [a, b, c]
    out.sort()
    for x in list_list:
        if out[-1] == len(x):
            return x

def run_milestone_key_checker_single(function, project_name, masters_list):
    '''

    function that runs this programme.

    :param function: The type of milestone you wish to analysis can be specified through choosing all_milestone_data_bulk,
    ap_p_milestone_data_bulk, or assurance_milestone_data_bulk functions, all available from engine_function import
    statement above.
    :param project_name: list of project to return data for
    :param masters_list: list of masters containing quarter information
    :return: excel wb.

    '''

    baseline_bc = bc_ref_stages([project_name], masters_list)
    baseline_list_index = master_baseline_index([project_name], masters_list, baseline_bc)

    p_current_milestones_data = function([project_name], masters_list[baseline_list_index[project_name][0]])
    p_last_milestones_data = function([project_name], masters_list[baseline_list_index[project_name][1]])
    p_oldest_milestones_data = function([project_name], masters_list[baseline_list_index[project_name][2]])

    run = check_m_keys_in_excel_single(project_name, p_current_milestones_data, p_last_milestones_data,
                                       p_oldest_milestones_data)

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
one_project_list = ['Commercial Vehicle Services (CVS)']

'''TWO. Specify date after which project milestones should be returned. NOTE: Python date format is (YYYY,MM,DD)'''
start_date = datetime.date(2019, 6, 1)

'''THREE. the following for statement prompts the programme to run. 
step one - place the list of projects chosen in step three at the end of the for statement. i.e. for project_name in [here] 
step two - chose the variables required for the run_milestone_comparator_single function. The first argument in relation
to which milestone data is to be analysed will normally be the only change. 
step three - provide relevant file path to document output. Changing the quarter stamp info as necessary. Note keep {} in 
file name as this is where the project name is recorded in the file title'''

for project_name in project_quarter_list:
    print('Doing milestone key name checking for ' + str(project_name))
    wb = run_milestone_key_checker_single(all_milestone_data_bulk, project_name, list_of_masters_all)
    wb.save('C:\\Users\\Standalone\\general\\masters folder\\project_milestones\\'
            'q2_1920_{}_milestone_keys_checker.xlsx'.format(project_name))