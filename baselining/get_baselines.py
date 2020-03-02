'''Throw away code for compiling bl information on projects, used to commission and collect further data from projects
output is excel wb with project bl data on seperate ws.
Could be of future use'''

from openpyxl import Workbook
from analysis.data import list_of_masters_all, latest_quarter_project_names, root_path, baseline_bc_stamp
from analysis.engine_functions import get_quarter_stamp, baseline_information

def get_baseline(project_name_list, baseline_data, comp_baseline_list):
    wb = Workbook()

    for i, name in enumerate(project_name_list):
        '''worksheet is created for each project'''
        ws = wb.create_sheet(name, i)  # creating worksheets
        ws.title = name  # title of worksheet

        for data in (baseline_data[name]):
            column_index = data[2]
            bc_stage = data[0]
            ws.cell(row=2, column=column_index+2, value=bc_stage)

        for x, baseline_type in enumerate(comp_baseline_list):
            for albm in (baseline_type[name]):
                column_index = albm[2]
                bc_stage = albm[0]
                ws.cell(row=x+3, column=column_index+2, value=bc_stage)

        quarter_labels = get_quarter_stamp(list_of_masters_all)
        for l, label in enumerate(quarter_labels):
            ws.cell(row=1, column=l+2, value=label)

        for n, key in enumerate(baseline_name_list[1:]):
            ws.cell(row=n+4, column=1, value=key)

        ws.cell(row=1, column=1, value='Quarter')
        ws.cell(row=2, column=1, value='IPDC bc approval')
        ws.cell(row=3, column=1, value='Re-baselined this quarter')


    return wb

baseline_name_list = ['this quarter',
                      'ALB milestones',
                      'ALB cost',
                      'ALB benefits',
                      'IPDC milestones',
                      'IPDC cost',
                      'IPDC benefits',
                      'HMT milestones',
                      'HMT cost',
                      'HMT benefits']

all_baselines_list = []
for x in baseline_name_list:
    baselines = baseline_information(latest_quarter_project_names, list_of_masters_all, x)
    all_baselines_list.append(baselines)

'''Running the programme'''
'''output one - all data'''
run = get_baseline(latest_quarter_project_names, baseline_bc_stamp, all_baselines_list)

'''Specify name of the output document here'''
run.save(root_path/'output/baseline_info.xlsx')
