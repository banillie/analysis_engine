'''

Programme for pulling out single data point across chosen number of quarters.

It outputs a workbook with some conditional formatting to show, 1) changes in reported data - highlighted by salmon
pink background, 2)when projects were not reporting cell is grey, 3) the relevant colour of the data represents a rag
status

Follow instruction as set out below are provided


'''

from openpyxl import Workbook
from bcompiler.utils import project_data_from_master
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
import random
from data import q1_1920, one_quarter_dict_list, bespoke_group_dict_list, list_of_dicts_all
from engine_functions import all_milestone_data_bulk

def data_return(dict_list, project_list, data_key):
    ''' places all (non milestone) data of interest into excel file output '''

    salmon_fill = PatternFill(start_color='ff8080', end_color='ff8080', fill_type='solid')
    # red_text = Font(color="FF0000") #currently not in use

    wb = Workbook()
    ws = wb.active

    '''lists project names in ws'''
    for x in range(0, len(project_list)):
        try:
            ws.cell(row=x + 2, column=1).value = dict_list[0][project_list[x]]['DfT Group']
        except KeyError:
            pass
        ws.cell(row=x + 2, column=2, value=project_list[x])

    '''project data into ws'''
    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=2).value
        print(project_name)
        col_start = 3
        for i, dictionary in enumerate(dict_list):
            if project_name in dictionary:
                ws.cell(row=row_num, column=col_start).value = dictionary[project_name][data_key]
                if dictionary[project_name][data_key] == None:
                    ws.cell(row=row_num, column=col_start).value = 'None'
                try:
                    if dict_list[i+1][project_name][data_key] != dictionary[project_name][data_key]:
                        ws.cell(row=row_num, column=col_start).fill = salmon_fill
                except (IndexError, KeyError):
                    pass
                col_start += 1
            else:
                ws.cell(row=row_num, column=col_start).value = 'Not reporting'
                col_start += 1

    '''quarter tag / meta data into ws'''
    quarter_labels = get_quarter_stamp(dict_list)
    ws.cell(row=1, column=1, value='Group')
    ws.cell(row=1, column=2, value='Project')
    for i, label in enumerate(quarter_labels):
        ws.cell(row=1, column=i + 3, value=label)

    conditional_formatting(ws)  # apply conditional formatting

    return wb

def milestone_data_return(dict_list, project_list, data_key):
    ''' places all (non milestone) data of interest into excel file output '''

    salmon_fill = PatternFill(start_color='ff8080', end_color='ff8080', fill_type='solid')
    # red_text = Font(color="FF0000") #currently not in use

    wb = Workbook()
    ws = wb.active

    '''lists project names in ws'''
    for x in range(0, len(project_list)):
        try:
            ws.cell(row=x + 2, column=1).value = dict_list[0][project_list[x]]['DfT Group']
        except KeyError:
            pass
        ws.cell(row=x + 2, column=2, value=project_list[x])

    '''project data into ws'''
    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=2).value
        print(project_name)
        col_start = 3
        for i, dictionary in enumerate(dict_list):
            if project_name in dictionary:

                milestone_dict = all_milestone_data_bulk([project_name], dictionary)
                #print(milestone_dict)

                try:
                    ws.cell(row=row_num, column=col_start).value = tuple(milestone_dict[project_name][data_key])[0]

                    if tuple(milestone_dict[project_name][data_key])[0] is None:
                        ws.cell(row=row_num, column=col_start).value = 'None'
                except KeyError:
                    ws.cell(row=row_num, column=col_start).value = 'None'

                try:

                    last_milestone_dict = all_milestone_data_bulk([project_name], dict_list[i + 1])

                    if tuple(last_milestone_dict[project_name][data_key])[0] != \
                            tuple(milestone_dict[project_name][data_key])[0]:
                        ws.cell(row=row_num, column=col_start).fill = salmon_fill
                except (IndexError, KeyError):
                    pass
                col_start += 1
            else:
                ws.cell(row=row_num, column=col_start).value = 'Not reporting'
                col_start += 1

    '''quarter tag / meta data into ws'''
    quarter_labels = get_quarter_stamp(dict_list)
    ws.cell(row=1, column=1, value='Group')
    ws.cell(row=1, column=2, value='Project')
    for i, label in enumerate(quarter_labels):
        ws.cell(row=1, column=i + 3, value=label)

    conditional_formatting(ws)  # apply conditional formatting

    return wb

def conditional_formatting(worksheet):

    '''function for applying rag rating conditional formatting colouring if required'''

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

def get_all_project_names(dict_list):

    '''returns list of all projects across multiple dictionaries'''
    output_list = []
    for dict in dict_list:
        for name in dict:
            if name not in output_list:
                output_list.append(name)

    return output_list

def get_quarter_stamp(dict_list):

    '''used to specify the quarter being reported'''

    output_list = []
    for dict in dict_list:
        proj_name = random.choice(list(dict.keys()))
        quarter_stamp = dict[proj_name]['Reporting period (GMPP - Snapshot Date)']
        output_list.append(quarter_stamp)

    return output_list


''' RUNNING PROGRAMME '''

'''Note that the all master data is taken from the data file'''

''' ONE. Set relevant list of projects. This needs to be done in accordance with the data you are working with via the
 data.py file '''
one_quarter_list = list(q1_1920.keys())
combined_quarters_list = get_all_project_names(list_of_dicts_all)
specific_project_list = [] # opportunity to provide manual list of projects

'''TWO. Set data of interest. there are two options here. hash out whichever option you are not using'''

'''option one - non-milestone data'''
#data_interest = 'DfT Group'

'''option two - milestone data'''
milestone_data_interest = 'Project End Date'

'''THREE. Run the programme'''

'''option one - run the data_return function for all non-milestone data'''
#run = data_return(list_of_dicts_all, combined_quarters_list, data_interest)

'''option two - run the milestone_data_return for all milestone data'''
run = milestone_data_return(list_of_dicts_all, one_quarter_list, milestone_data_interest)

'''FOUR. specify the file path and name of the output document'''
run.save('C:\\Users\\Standalone\\general\\project_end_date.xlsx')
