'''

This programme creates a master spreadsheet for commissioning templates.

Documents required to run the programme are set out below. The latest versions of these should be taken from TiME
and saved onto laptops in the file paths at the bottom of the programme.

Documents required are:
1) the internal commission datamap (make sure you have the latest/final version).
2) latest quarter DfT_master spreadsheet. (Note in the commission the latest quarter data is used to both populate the
commission template in active reporting fields, as well as provide a record of what is reported last quarter.)

2) and 3) above are taken from the data module

'''

from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from analysis.data import q1_1920


def create_master(workbook, latest_q_data):
    ws = workbook.active
    type_list_2 = ['RDEL', 'CDEL', 'Non-Gov', 'Income', 'BEN'] # list of cost/ben types. used to remove none value entries

    red_text = Font(color="00fc2525")

    for i, name in enumerate(list(latest_q_data.keys())):
        print(name)
        ws.cell(row=1, column=2+i).value = name  # place project names in file

        # for loop for placing data into the worksheet
        for row_num in range(2, ws.max_row+1):
            key = ws.cell(row=row_num, column=1).value

            if key in latest_q_data[name].keys():
                ws.cell(row=row_num, column=2+i).value = latest_q_data[name][key]
            elif key not in latest_q_data[name].keys():
                altered_lst_quarter = key.replace("lst qrt ", "")
                ws.cell(row=row_num, column=2 + i).value = latest_q_data[name][altered_lst_quarter]
                ws.cell(row=row_num, column=2 + i).font = red_text
            else:
                ws.cell(row=row_num, column=40).value = 'Couldnt match'
                #print(key)

                # this section of the code ensures that all financial costs / benefit forecasts have a zero
                for cost_type in type_list_2:
                    if cost_type in key:
                        try:
                            if latest_q_data[name][key] is None:
                                ws.cell(row=row_num, column=2 + i).value = 0
                        except KeyError:
                            pass

    return workbook


latest_dm = load_workbook("C:\\Users\\Standalone\\general\\masters folder\\commission\\Q2_1920_commission_testing.xlsx")
# 1) place file path to gmpp data map here

run = create_master(latest_dm, q1_1920)

run.save("C:\\Users\\Standalone\\general\\test_2.xlsx")
# Place file path for whether you want file to be saved and what you'd like it to be called here