"""
Transfers MilestoneData object into an excel wb. Wb includes calculation
of time differences between milestone dates at current, last and
baseline quarter.
"""

from openpyxl import Workbook
from data_mgmt.data import MilestoneData, MilestoneChartData, \
    Masters, Projects, master_data_list, root_path, blue_line_date, \
    abbreviations, CombinedData

def put_into_wb_all(milestone_data_object):
    wb = Workbook()
    ws = wb.active

    row_num = 2
    for project_name in milestone_data_object.project_current.keys():
        for i, milestone in enumerate(milestone_data_object.project_current_p_project[project_name].keys()):
            ws.cell(row=row_num + i, column=1).value = project_name
            ws.cell(row=row_num + i, column=2).value = milestone
            try:
                milestone_date = tuple(milestone_data_object.project_current_p_project[project_name][milestone])[0]
                ws.cell(row=row_num + i, column=3).value = milestone_date
                ws.cell(row=row_num + i, column=3).number_format = 'dd/mm/yy'
            except KeyError:
                ws.cell(row=row_num + i, column=3).value = ''

            try:
                last_date = tuple(milestone_data_object.project_last_p_project[project_name][milestone])[0]
                ws.cell(row=row_num + i, column=4).value = (milestone_date - last_date).days
            except (KeyError, TypeError):
                ws.cell(row=row_num + i, column=4).value = ''

            try:
                baseline_date = tuple(milestone_data_object.project_baseline_p_project[project_name][milestone])[0]
                ws.cell(row=row_num + i, column=5).value = (milestone_date - baseline_date).days
            except (KeyError, TypeError):
                ws.cell(row=row_num + i, column=5).value = ''

            try:
                notes = milestone_data_object.project_current_p_project[project_name][milestone][milestone_date]
                ws.cell(row=row_num + i, column=7).value = notes
            except (IndexError, KeyError):
                ws.cell(row=row_num + i, column=7).value = ''

        row_num = row_num + len(milestone_data_object.project_current_p_project[project_name].keys())

    ws.cell(row=1, column=1).value = 'Project'
    ws.cell(row=1, column=2).value = 'Milestone'
    ws.cell(row=1, column=3).value = 'Date'
    ws.cell(row=1, column=4).value = '3/m change'
    ws.cell(row=1, column=5).value = 'Baseline change (current)'
   # ws.cell(row=1, column=6).value = 'Baseline change (last)'
    ws.cell(row=1, column=7).value = 'Notes'

    return wb

mst = Masters(master_data_list, Projects.current_list)  # get master data and specify projects
mst.baseline_data('Re-baseline IPDC milestones')  # get baseline information of interest
milestone_data = MilestoneData(mst, abbreviations)  # get milestone data

run = put_into_wb_all(milestone_data)
run.save(root_path/"output/all_delivery_milestones_q1_20_21.xlsx")