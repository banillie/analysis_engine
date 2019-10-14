'''
programme that generate the information required for the dandelion chart. The output data is then sorted in excel and
placed into the dandelion graph wb.

Follow instructions below.

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.
'''

from openpyxl import Workbook
from analysis.data import q1_1920

def dandelion_data(master_data):
    '''
    Simple function that returns data required for the dandelion graph. Sorting done via excel.

    :param master_data: quarter master data
    :return: excel wb
    '''

    wb = Workbook()
    ws = wb.active

    for i, project_name in enumerate(master_data.projects):
        ws.cell(row=2 + i, column=1).value = master_data.data[project_name]['DfT Group']
        ws.cell(row=2 + i, column=2).value = project_name
        ws.cell(row=2 + i, column=3).value = master_data.data[project_name]['Total Forecast']

    ws.cell(row=1, column=1).value = 'Group'
    ws.cell(row=1, column=2).value = 'Project Name'
    ws.cell(row=1, column=3).value = 'WLC (forecast)'

    return wb

'''  RUNNING PROGRAMME '''

'''Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.

Note. much of the work required for the final output is done in excel. Refer to separate gist hub guidance'''

'''ONE. place the master quarter data of interest into the dandelion data function and specify the file path for where 
the output excel file should be saved'''

output = dandelion_data(q1_1920)
output.save('C:\\Users\\Standalone\\general\\masters folder\\dandelion\\testing.xlsx')