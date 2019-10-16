'''
Programme for creating the gmpp address reference for gmpp master data. might be useful for the future.
'''

from datamaps.api import project_data_from_master
from openpyxl import Workbook

gmpp_dm = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\gmpp_reporting\\'
                                   'gmpp_datamaps\\gmpp_datamap_q2_1920.xlsx', 2, 2019)

def create_address(dm):
    a = list(dm.data['template_sheet'].values())

    b = list(dm.data['cell_reference'].values())

    output = []
    for i, name in enumerate(a):
        new_string = str(name) + "'!" + str(b[i])
        output.append(new_string)

    wb = Workbook()
    ws = wb.active

    for i, name in enumerate(output):
        ws.cell(row=1 + i, column=1).value = name

    return wb

run = create_address(gmpp_dm)

run.save('C:\\Users\\Standalone\\general\\masters folder\\gmpp_reporting\\gmpp_datamaps\\gmpp_data_q2_1920.xlsx')