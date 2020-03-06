'''
Creates the 'address reference' used by IPA for gmpp reporting from the DfT internal cell_reference and template_
sheet.

Output goes into wb.

probably throw away code
'''

from datamaps.api import project_data_from_master
from openpyxl import Workbook
from analysis.data import root_path

gmpp_dm = project_data_from_master(root_path/'input/gmpp_master_dm.xlsx', 2, 2019)

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

run.save(root_path/'output/test.xlsx')
