'''

This programme creates a master spreadsheet to share with IPA for gmpp reporting. The 'master' print out is then
shared with the IPA which runs an excel macro to populate individual gmpp reporting templates.

Documents required to run the programme are set out below. The latest versions of these should be taken from TiME
and saved onto laptops in the file paths at the bottom of the programme.

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

#import datetime
from datamaps.api import project_data_from_master
from openpyxl import Workbook, load_workbook
#from collections import OrderedDict
#from openpyxl.utils import column_index_from_string
from openpyxl.styles import Border, Color, Font, PatternFill


def create_master(gmpp_wb, latest_data, last_gmpp):
    ws = gmpp_wb.active

    type_list = ['RDEL', 'CDEL', 'Non-Gov', 'Income'] # list of cost types. used to amend Hs2 data
    type_list_2 = ['RDEL', 'CDEL', 'Non-Gov', 'Income', 'BEN'] # list of cost/ben types. used to remove none value entries

    red_text = Font(color="00fc2525")

    # this section filters out only gmpp project names. Subsequent list is then used to populate ws
    gmpp_project_names = []
    for name in (list(latest_data.keys())):
        #print(name)
        if latest_data[name]['GMPP - IPA DCA last quarter'] != None:
            #print(name)
            gmpp_project_names.append(name)

    for i, name in enumerate(gmpp_project_names):
        print(name)
        ws.cell(row=1, column=5+i).value = name  # place project names in file

        # for loop for placing data into the worksheet
        for row_num in range(2, ws.max_row+1):
            key = ws.cell(row=row_num, column=1).value
            # this loop places all latest raw data into the worksheet
            if key in latest_data[name].keys():
                ws.cell(row=row_num, column=5+i).value = latest_data[name][key]
            # elif key not in latest_data[name].keys():
            #     print(key)

                # this section of the code ensures that all financial costs / benefit forecasts have a zero
                for cost_type in type_list_2:
                    if cost_type in key:
                        if latest_data[name][key] is None:
                            ws.cell(row=row_num, column=5 + i).value = 0

            # # this section handles some easily automated tweaks to data to meet IPA data structures for non-static data
            # if key == 'Project/Programme Name':
            #     ws.cell(row=row_num, column=11+i).value = name
            # #if key == 'FD Sign-Off':
            # #    ws.cell(row=row_num, column=11+i).value = None
            # #if key == 'New PD - If \'other\' please specify':
            # #    ws.cell(row=row_num, column=11+i).value = None
            # #if key == 'Person Completing this return: Email Address':
            # #    ws.cell(row=row_num, column=11+i).value = 'robert.green@dft.gov.uk'
            # if key == 'Dept Single Point of Contact (SPOC)':
            #     ws.cell(row=row_num, column=11+i).value = 'Robert Green'          # HARDCODED
            # if key == 'SPOC Email Address':
            #     ws.cell(row=row_num, column=11+i).value = 'robert.green@dft.gov.uk'   # HARDCODED
            # if key == 'Snapshot':
            #     ws.cell(row=row_num, column=11+i).value = '3\1-Dec-18'      # HARDCODED. Change each quarter.


            # if key == 'SRO First Name':
            #     email = latest_data[name]['SRO Email']
            #     email_1 = email.split("@")[0]
            #     firstname = email.split(".")[0]
            #     ws.cell(row=row_num, column=11+i).value = firstname
            # if key == 'SRO Last Name':
            #     email = latest_data[name]['SRO Email']
            #     email_1 = email.split("@")[0]
            #     surname = email_1.split(".")[1]
            #     ws.cell(row=row_num, column=11+i).value = surname
            # if key == 'PD First Name':
            #     email = latest_data[name]['PD Email']
            #     email_1 = email.split("@")[0]
            #     firstname = email_1.split(".")[0]
            #     ws.cell(row=row_num, column=11+i).value = firstname
            # if key == 'PD Last Name':
            #     email = latest_data[name]['PD Email']
            #     email_1 = email.split("@")[0]
            #     surname = email_1.split(".")[1]
            #     ws.cell(row=row_num, column=11+i).value = surname
            #
            # if key == 'Project Cost Narrative':
            #     rdel = latest_data[name]['Project Costs Narrative RDEL']
            #     if rdel == None:
            #         rdel = ''
            #     cdel = latest_data[name]['Project Costs Narrative CDEL']
            #     if cdel == None:
            #         cdel = ''
            #     ws.cell(row=row_num, column=11+i).value = rdel + cdel
            #     #ws.cell(row=row_num, column=col_num).font = red_text
            #
            # '''this loop places all static gmpp specific information into worksheet. needs some further work
            # this can overwrite in uncontrolled way new data being put into sheet. have made an attempt to fix this
            # on list 116 below'''
            # if key not in latest_data[name].keys():
            #     if key in last_gmpp[name].keys():
            #         if key not in list_gmpp_static_keys:
            #             print(key)
            #             ws.cell(row=row_num, column=11 + i).value = last_gmpp[name][key]

    # this section handles HS2 data. placing old static data into the worksheet
    for i, name in enumerate(gmpp_project_names):
        if name == 'High Speed Rail Programme (HS2)':
            print('HS2 financial data has been amended')
            '''note minus 20 here. bug in the loop I haven't fixed yet. probably something to do with how data is
            recorded in DM'''
            for row_num in range(2, ws.max_row-20):
                key = ws.cell(row=row_num, column=1).value
                for cost_type in type_list:
                    try:
                        if cost_type in key:
                            ws.cell(row=row_num, column=5 + i).value = last_gmpp[name][key]
                            ws.cell(row=row_num, column=5 + i).font = red_text
                    except (KeyError, TypeError):
                        pass

    return gmpp_wb

# list_gmpp_static_keys = ['SRO Last Name', 'SRO First Name', 'PD Last Name', 'PD First Name', 'First Name',
#                          'Last Name', 'Project Costs Narrative']

latest_dm = load_workbook("C:\\Users\\Standalone\\general\\masters folder\\gmpp_reporting\\gmpp_datamaps\\"
                          "gmpp_datamap.xlsx")    # 1) place file path to gmpp data map here

dft_master = project_data_from_master("C:\\Users\\Standalone\\general\\masters folder\\core_data\\master_1_2019.xlsx")
# 2) place file path to latest quarter master above

old_hs2_master = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\gmpp_reporting\\2018_19\\'
                                          'Q3_1819\\gmpp_master_Q3.xlsx')
# 3) place file path to last quarter gmpp master here

run = create_master(latest_dm, dft_master, old_hs2_master)

run.save("C:\\Users\\Standalone\\general\\test.xlsx")
