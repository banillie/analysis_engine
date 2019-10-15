'''

Programme for pulling out single data point across chosen number of quarters.

It outputs a workbook, which shows:
1) reported data across multiple quarters
2) changes in reported data - highlighted by salmon pink background,
3) when projects were not reporting data -  grey out cell,
4) if a rag status is returned the colour of the rag status

Follow instruction as set out below are provided

'''

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from analysis.data import q1_1920, one_quarter_master_list, bespoke_group_masters_list, list_of_masters_all
from analysis.engine_functions import all_milestone_data_bulk, ap_p_milestone_data_bulk, assurance_milestone_data_bulk,\
    get_all_project_names, get_quarter_stamp

def data_return(masters_list, project_name_list, data_key):
    '''
    places all (non-milestone) data of interest into excel file output

    masters_list: list of masters containing quarter information
    project_name_list: list of project to return data for
    data_key: the data key of interest

    '''

    salmon_fill = PatternFill(start_color='ff8080', end_color='ff8080', fill_type='solid')
    # red_text = Font(color="FF0000") #currently not in use

    wb = Workbook()
    ws = wb.active

    '''lists project names in ws'''
    for x in range(0, len(project_name_list)):
        try:
            ws.cell(row=x + 2, column=1).value = masters_list[0].data[project_name_list[x]]['DfT Group']
        except KeyError:
            pass
        ws.cell(row=x + 2, column=2, value=project_name_list[x])

    '''project data into ws'''
    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=2).value
        print(project_name)
        col_start = 3
        for i, master in enumerate(masters_list):
            if project_name in master.projects:
                try:
                    ws.cell(row=row_num, column=col_start).value = master.data[project_name][data_key]
                    if master.data[project_name][data_key] == None:
                        ws.cell(row=row_num, column=col_start).value = 'None'
                    try:
                        if masters_list[i+1].data[project_name][data_key] != master.data[project_name][data_key]:
                            ws.cell(row=row_num, column=col_start).fill = salmon_fill
                    except (IndexError, KeyError):
                        pass
                    col_start += 1
                except KeyError:
                    ws.cell(row=row_num, column=col_start).value = 'data key not collected'
            else:
                ws.cell(row=row_num, column=col_start).value = 'Not reporting'
                col_start += 1

    '''quarter tag / meta data into ws'''
    quarter_labels = get_quarter_stamp(masters_list)
    ws.cell(row=1, column=1, value='Group')
    ws.cell(row=1, column=2, value='Project')
    for i, label in enumerate(quarter_labels):
        ws.cell(row=1, column=i + 3, value=label)

    #conditional_formatting(ws)  # apply conditional formatting

    return wb

def milestone_data_return(masters_list, project_name_list, data_key):
    ''' places all milestone data of interest into excel file output

    master_list: list of masters containing quarter information
    project_name_list: list of project to return data for
    data_key: the data key of interest
    '''

    salmon_fill = PatternFill(start_color='ff8080', end_color='ff8080', fill_type='solid')
    # red_text = Font(color="FF0000") #currently not in use

    wb = Workbook()
    ws = wb.active

    '''lists project names in ws'''
    for x in range(0, len(project_name_list)):
        try:
            ws.cell(row=x + 2, column=1).value = masters_list[0].data[project_name_list[x]]['DfT Group']
        except KeyError:
            pass
        ws.cell(row=x + 2, column=2, value=project_name_list[x])

    '''project data into ws'''
    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=2).value
        print(project_name)
        col_start = 3
        for i, master in enumerate(masters_list):
            if project_name in master.projects:

                milestone_data = all_milestone_data_bulk([project_name], master)

                try:
                    ws.cell(row=row_num, column=col_start).value = tuple(milestone_data[project_name][data_key])[0]

                    if tuple(milestone_data[project_name][data_key])[0] is None:
                        ws.cell(row=row_num, column=col_start).value = 'None'
                except KeyError:
                    ws.cell(row=row_num, column=col_start).value = 'None'

                try:

                    last_milestone_data = all_milestone_data_bulk([project_name], masters_list[i + 1])

                    if tuple(last_milestone_data[project_name][data_key])[0] != \
                            tuple(milestone_data[project_name][data_key])[0]:
                        ws.cell(row=row_num, column=col_start).fill = salmon_fill
                except (IndexError, KeyError):
                    pass
                col_start += 1
            else:
                ws.cell(row=row_num, column=col_start).value = 'Not reporting'
                col_start += 1

    '''quarter tag / meta data into ws'''
    quarter_labels = get_quarter_stamp(masters_list)
    ws.cell(row=1, column=1, value='Group')
    ws.cell(row=1, column=2, value='Project')
    for i, label in enumerate(quarter_labels):
        ws.cell(row=1, column=i + 3, value=label)

    #conditional_formatting(ws)  # apply conditional formatting

    return wb

def conditional_formatting(worksheet):

    '''
    Function for applying rag rating conditional formatting colouring.

    This doesn't always need to be applied. ToDo: Need to think/design and way of placing this into instructions.

    '''

    ag_text = Font(color="000000")
    ag_fill = PatternFill(bgColor="00a5b700")
    dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber/Green", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber/Green",A1)))']
    worksheet.conditional_formatting.add('A1:X80', rule)

    ar_text = Font(color="000000")
    ar_fill = PatternFill(bgColor="00f97b31")
    dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber/Red", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber/Red",A1)))']
    worksheet.conditional_formatting.add('A1:X80', rule)

    red_text = Font(color="000000")
    red_fill = PatternFill(bgColor="00fc2525")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="Red", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Red",A1)))']
    worksheet.conditional_formatting.add('A1:X80', rule)

    green_text = Font(color="000000")
    green_fill = PatternFill(bgColor="0017960c")
    dxf = DifferentialStyle(font=green_text, fill=green_fill)
    rule = Rule(type="containsText", operator="containsText", text="Green", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Green",A1)))']
    worksheet.conditional_formatting.add('A1:X80', rule)

    amber_text = Font(color="000000")
    amber_fill = PatternFill(bgColor="00fce553")
    dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber",A1)))']
    worksheet.conditional_formatting.add('A1:X80', rule)

    grey_text = Font(color="f0f0f0")
    grey_fill = PatternFill(bgColor="f0f0f0")
    dxf = DifferentialStyle(font=grey_text, fill=grey_fill)
    rule = Rule(type="containsText", operator="containsText", text="Not reporting", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Not reporting",A1)))']
    worksheet.conditional_formatting.add('A1:X80', rule)

    # highlighting new projects
    red_text = Font(color="000000")
    white_fill = PatternFill(bgColor="000000")
    dxf = DifferentialStyle(font=red_text, fill=white_fill)
    rule = Rule(type="containsText", operator="containsText", text="NEW", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("NEW",A1)))']
    worksheet.conditional_formatting.add('A1:X80', rule)

    return worksheet

''' RUNNING PROGRAMME '''

'''Note that the all master data is taken from the data file. Make sure that this is up to date and that all relevant
  data is being imported'''

''' ONE. Set relevant list of projects. This needs to be done in accordance with the data you are working with via the
 data.py file '''
one_quarter_list = q1_1920.projects
combined_quarters_list = get_all_project_names(list_of_masters_all)
specific_project_list = [] # opportunity to provide manual list of projects

'''TWO. Set data of interest. there are two options here. hash out whichever option you are not using'''

'''option one - non-milestone data'''
data_interest = 'SRO Full Name'

'''option two - milestone data'''
#milestone_data_interest = 'Project End Date'

'''THREE. Run the programme'''

'''option one - run the data_return function for all non-milestone data'''
run = data_return(list_of_masters_all, one_quarter_list, data_interest)

'''option two - run the milestone_data_return for all milestone data'''
#run = milestone_data_return(list_of_masters_all, one_quarter_list, milestone_data_interest)

'''FOUR. specify the file path and name of the output document'''
run.save('C:\\Users\\Standalone\\general\\sros_for_projects.xlsx')
