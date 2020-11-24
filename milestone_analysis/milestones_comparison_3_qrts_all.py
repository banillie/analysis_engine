"""
Transfers MilestoneData object into an excel wb. Wb includes calculation
of time differences between milestone dates at current, last and
baseline quarter.
"""

from openpyxl import Workbook
from data_mgmt.data import MilestoneData, Master, Projects, \
    get_master_data, abbreviations, root_path, get_current_project_names

def put_into_wb_all(milestone_data_object):
    wb = Workbook()
    ws = wb.active

    row_num = 2
    for project_name in milestone_data_object.current.keys():
        for i, milestone in enumerate(milestone_data_object.current[project_name].keys()):
            ws.cell(row=row_num + i, column=1).value = project_name
            ws.cell(row=row_num + i, column=2).value = milestone
            try:
                milestone_date = tuple(milestone_data_object.current[project_name][milestone])[0]
                ws.cell(row=row_num + i, column=3).value = milestone_date
                ws.cell(row=row_num + i, column=3).number_format = 'dd/mm/yy'
            except KeyError:
                ws.cell(row=row_num + i, column=3).value = ''

            try:
                last_date = tuple(milestone_data_object.last_quarter[project_name][milestone])[0]
                ws.cell(row=row_num + i, column=4).value = (milestone_date - last_date).days
            except (KeyError, TypeError):
                ws.cell(row=row_num + i, column=4).value = ''

            try:
                baseline_date = tuple(milestone_data_object.baseline[project_name][milestone])[0]
                ws.cell(row=row_num + i, column=5).value = (milestone_date - baseline_date).days
            except (KeyError, TypeError):
                ws.cell(row=row_num + i, column=5).value = ''

            try:
                notes = milestone_data_object.current[project_name][milestone][milestone_date]
                ws.cell(row=row_num + i, column=7).value = notes
            except (IndexError, KeyError):
                ws.cell(row=row_num + i, column=7).value = ''

        row_num = row_num + len(milestone_data_object.current[project_name].keys())

    ws.cell(row=1, column=1).value = 'Project'
    ws.cell(row=1, column=2).value = 'Milestone'
    ws.cell(row=1, column=3).value = 'Date'
    ws.cell(row=1, column=4).value = '3/m change'
    ws.cell(row=1, column=5).value = 'Baseline change (current)'
   # ws.cell(row=1, column=6).value = 'Baseline change (last)'
    ws.cell(row=1, column=7).value = 'Notes'

    return wb

master_data = get_master_data()
current_project_name_list = get_current_project_names()

mst = Master(master_data, current_project_name_list)  # get Master object and specify projects of interest
mst.baseline_data('Re-baseline IPDC milestones')  # place baseline information of interest into master object
milestone_data = MilestoneData(mst, abbreviations)  # create MilestoneData object
milestone_data.get_milestones('Delivery')  # place type of milestone data of interest into MilestoneData object

run = put_into_wb_all(milestone_data)
run.save(root_path/"output/milestone_data_output.xlsx")