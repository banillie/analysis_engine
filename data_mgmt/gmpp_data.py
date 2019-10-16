'''

This programme creates a master spreadsheet to share with IPA for gmpp reporting. The 'master' print out is then
shared with the IPA which runs an excel macro to populate individual gmpp reporting templates.

Documents required to run the programme are set out below.

Documents required are:
1) the gmpp datamap (make sure you have the latest/final version).
2) The latest quarter DfT_master spreadsheet (i.e. the quarter that is being reported).
3) Last quarters gmpp master document (This document is used to provide static information). NOTE. Still needs to do a
little testing to ensure this part of project is working properly. Reported issue about project cost narratives. I think
this is now fixed, but need more testing. need to consider any impact on Hs2 data.

IMPORTANT to note:
- HS2 data has to be handled carefully. The financial data reported should be static/unchanged. This programme
handles this issue. The data is amended a placed into the master as red text. However the output master should be
manually checked to verify that the data is red.

'''

from openpyxl import load_workbook
from openpyxl.styles import Font
from analysis.data import q2_1920
from analysis.engine_functions import filter_gmpp


def create_master(gmpp_wb, master_data):
    ws = gmpp_wb.active

    type_list = ['RDEL', 'CDEL', 'Non-Gov', 'Income'] # list of cost types. used to amend Hs2 data
    type_list_2 = ['RDEL', 'CDEL', 'Non-Gov', 'Income', 'BEN'] # list of cost/ben types. used to remove none value entries

    red_text = Font(color="00fc2525")

    # this section filters out only gmpp project names. Subsequent list is then used to populate ws
    gmpp_project_names = filter_gmpp(master_data)

    for i, project_name in enumerate(gmpp_project_names):
        print(project_name)
        ws.cell(row=1, column=6+i).value = project_name  # place project names in file

        # for loop for placing data into the worksheet
        for row_num in range(2, ws.max_row+1):
            key = ws.cell(row=row_num, column=1).value
            # this loop places all latest raw data into the worksheet
            if key in master_data.data[project_name].keys():
                ws.cell(row=row_num, column=6+i).value = master_data.data[project_name][key]
            # elif key not in latest_data[name].keys():
            #     print(key)

                # this section of the code ensures that all financial costs / benefit forecasts have a zero
                for cost_type in type_list_2:
                    if cost_type in key:
                        if master_data.data[project_name][key] is None:
                            ws.cell(row=row_num, column=6 + i).value = 0

    return gmpp_wb

# list_gmpp_static_keys = ['SRO Last Name', 'SRO First Name', 'PD Last Name', 'PD First Name', 'First Name',
#                          'Last Name', 'Project Costs Narrative']

latest_dm = load_workbook("C:\\Users\\Standalone\\general\\masters folder\\gmpp_reporting\\gmpp_datamaps\\"
                          "gmpp_datamap_q2_1920.xlsx")

run = create_master(latest_dm, q2_1920)

run.save("C:\\Users\\Standalone\\general\\gmpp_dataset_q2_1920.xlsx")
