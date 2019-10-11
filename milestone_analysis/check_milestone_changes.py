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
from analysis.engine_functions import all_milestone_data_bulk, ap_p_milestone_data_bulk, assurance_milestone_data_bulk, \
    project_time_difference, bc_ref_stages, master_baseline_index, filter_project_group
from analysis.data import q2_1920, list_of_masters_all

'''Function that checks whether reported milestone keys have changed between quarters'''
def check_m_keys_in_excel_single(name, t_dict_one, t_dict_two, t_dict_three):
    wb = Workbook()
    ws = wb.active
    red_text = Font(color="00fc2525")

    row_num = 2

    one = list(t_dict_one[name].keys())
    [x for x in one if x is not None].sort()
    two = list(t_dict_two[name].keys())
    [x for x in two if x is not None].sort()
    three = list(t_dict_three[name].keys())
    [x for x in three if x is not None].sort()

    long = longest_list(one, two, three)
    for i in range(0, len(long)):
        ws.cell(row=row_num + i, column=1).value = name
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
    for i, name in enumerate(row_heading_list):
        ws.cell(row=1, column=i+1).value = name

    column_ltr_list = ['A', 'B', 'C', 'D', 'E']
    for ltr in (column_ltr_list):
        ws.column_dimensions[ltr].width = 40

    return wb

'''helper function for check_m_keys_in_excle'''
def longest_list(one, two, three):
    list_list = [one, two, three]
    a = len(one)
    b = len(two)
    c = len(three)

    out = [a,b,c]
    out.sort()
    for x in list_list:
        if out[-1] == len(x):
            return x


def run_milestone_key_checker_single(function, proj_name, q_masters_dict_list, baseline_chain_dict):
    p_current_milestones_dict = function([proj_name], q_masters_dict_list[baseline_chain_dict[proj_name][0]])
    p_last_milestones_dict = function([proj_name], q_masters_dict_list[baseline_chain_dict[proj_name][1]])
    p_oldest_milestones_dict = function([proj_name], q_masters_dict_list[baseline_chain_dict[proj_name][2]])

    run = check_m_keys_in_excel_single(proj_name, p_current_milestones_dict, p_last_milestones_dict, p_oldest_milestones_dict)

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

'''TWO. Specify date after which project milestones should be returned. NOTE: Python date format is (YYYY,MM,DD)'''
start_date = datetime.date(2019, 6, 1)
me))