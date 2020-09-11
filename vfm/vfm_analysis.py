from openpyxl import Workbook
from data_mgmt.data import get_master_data, get_current_project_names, root_path


def compile_data(masters, project_name_list):
    wb = Workbook()
    ws = wb.active

    for row, project in enumerate(project_name_list):
        i = row + 2  # has to start with row 1 in excel
        ws.cell(row=i, column=1).value = masters[0].data[project]['DfT Group']
        ws.cell(row=i, column=2).value = project
        ws.cell(row=i, column=3).value = masters[0].data[project]['NPV for all projects ' \
                                                                  'and NPV for programmes if available']
        try:
            ws.cell(row=i, column=4).value = masters[1].data[project]['NPV for all projects ' \
                                                                      'and NPV for programmes if available']
        except KeyError:
            ws.cell(row=i, column=4).value = 'Not reporting'
        ws.cell(row=i, column=5).value = masters[0].data[project]['Adjusted Benefits Cost Ratio (BCR)']
        try:
            ws.cell(row=i, column=6).value = masters[1].data[project]['Adjusted Benefits Cost Ratio (BCR)']
        except KeyError:
            ws.cell(row=i, column=6).value = 'Not reporting'
        ws.cell(row=i, column=7).value = masters[0].data[project]['Initial Benefits Cost Ratio (BCR)']
        try:
            ws.cell(row=i, column=8).value = masters[1].data[project]['Initial Benefits Cost Ratio (BCR)']
        except KeyError:
            ws.cell(row=i, column=8).value = 'Not reporting'
        ws.cell(row=i, column=9).value = masters[0].data[project]['VfM Category single entry']
        try:
            ws.cell(row=i, column=10).value = masters[1].data[project]['VfM Category single entry']
        except KeyError:
            ws.cell(row=i, column=10).value = 'Not reporting'
        ws.cell(row=i, column=11).value = masters[0].data[project]['Present Value Cost (PVC)']
        try:
            ws.cell(row=i, column=12).value = masters[1].data[project]['Present Value Cost (PVC)']
        except KeyError:
            ws.cell(row=i, column=12).value = 'Not reporting'
        ws.cell(row=i, column=13).value = masters[0].data[project]['Present Value Benefit (PVB)']
        try:
            ws.cell(row=i, column=14).value = masters[1].data[project]['Present Value Benefit (PVB)']
        except KeyError:
            ws.cell(row=i, column=14).value = 'Not reporting'

    ws.cell(row=1, column=1).value = 'DfT Group'
    ws.cell(row=1, column=2).value = 'Project'
    ws.cell(row=1, column=3).value = 'NPV'
    ws.cell(row=1, column=4).value = 'NPV lst qrt'
    ws.cell(row=1, column=5).value = 'Adjusted BCR'
    ws.cell(row=1, column=6).value = 'Adjusted BCR lst qrt'
    ws.cell(row=1, column=7).value = 'Initial BCR'
    try:
        ws.cell(row=i, column=8).value = masters[1].data[project]['Initial Benefits Cost Ratio (BCR)']
    except KeyError:
        ws.cell(row=i, column=8).value = 'Not reporting'
    ws.cell(row=i, column=9).value = masters[0].data[project]['VfM Category single entry']
    try:
        ws.cell(row=i, column=10).value = masters[1].data[project]['VfM Category single entry']
    except KeyError:
        ws.cell(row=i, column=10).value = 'Not reporting'
    ws.cell(row=i, column=11).value = masters[0].data[project]['Present Value Cost (PVC)']
    try:
        ws.cell(row=i, column=12).value = masters[1].data[project]['Present Value Cost (PVC)']
    except KeyError:
        ws.cell(row=i, column=12).value = 'Not reporting'
    ws.cell(row=i, column=13).value = masters[0].data[project]['Present Value Benefit (PVB)']
    try:
        ws.cell(row=i, column=14).value = masters[1].data[project]['Present Value Benefit (PVB)']
    except KeyError:
        ws.cell(row=i, column=14).value = 'Not reporting'

    return wb


master_data = get_master_data()
current_project_name_list = get_current_project_names()

run = compile_data(master_data, current_project_name_list)
run.save(root_path / "output/vfm_data_output.xlsx")
