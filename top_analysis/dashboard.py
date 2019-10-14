'''

Programme for creating an aggregate project dashboard

input documents:
1) Dashboard master document - this is an excel file. This is the dashboard, but all data fields left blank.
Note. If project data does not get placed into the correct part of the master, check that the project name is
consistent with the name in master data, because names need to be exactly the same for information to be exported.

output document:
1) Dashboard with all project data placed into dashboard and formatted correctly.

Instructions:
1) provide path to dashboard master
2) change bicc_date variable
3) provide path and specify file name for output document

Supplementary instructions:
These things need to be done to check and assure the data going into the dashboard. Use the other programmes available
for undertaking these tasks.
1) Check that project stage/last at BICC data is correct. This is done via the bc_stage_from_master and
bc_amended_to_master programmes.
2) insert into the master document last at / next at BICC project data. This is done via the bicc_dates_from_master and
bicc_dates_amended_to_master programmes.

Note that some manual adjustments need to be made to:
1) Project WLC totals e.g. Hs2 Phases
2) The last/next at BICC specification. e.g. Hs2 Prog should be changed to 'often'


NOTE - code should ideally be refactored in line with latest structure as sort of hacked together at mo to accommodate
the change in data keys (i.e. new template).

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.
'''

from openpyxl import load_workbook
import datetime
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, IconSet, FormatObject
from analysis.data import q2_1920, q1_1920
from analysis.engine_functions import all_milestone_data_bulk, convert_rag_text, concatenate_dates, up_or_down, \
    convert_bc_stage_text

def place_in_excel(master_one, master_two, wb):
    '''
    function that places all information into the master dashboard sheet
    :param master_one:
    :param master_two:
    :return:
    '''

    ws = wb.active

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        print(project_name)
        if project_name in master_one.projects:
            ws.cell(row=row_num, column=4).value = master_one.data[project_name]['Total Forecast']
            try:
                dca_now = master_one.data[project_name]['Departmental DCA']
                dca_past = master_two.data[project_name]['Departmental DCA']
                ws.cell(row=row_num, column=6).value = up_or_down(dca_now, dca_past)
            except KeyError:
                ws.cell(row=row_num, column=6).value = 'NEW'
            ws.cell(row=row_num, column=7).value = convert_rag_text(master_one.data[project_name]['Departmental DCA'])
            ws.cell(row=row_num, column=8).value = convert_rag_text(master_one.data[project_name]
                                                                    ['GMPP - IPA DCA last quarter'])
            ws.cell(row=row_num, column=9).value = convert_bc_stage_text(master_one.data[project_name]
                                                                         ['BICC approval point'])

            p_m_data = all_milestone_data_bulk([project_name], master_one)
            print(p_m_data)
            ws.cell(row=row_num, column=10).value = concatenate_dates\
                (list(p_m_data[project_name]['Start of Operation'])[0])
            ws.cell(row=row_num, column=11).value = concatenate_dates\
                (list(p_m_data[project_name]['Project End Date'])[0])
            ws.cell(row=row_num, column=12).value = convert_rag_text(master_one.data[project_name]
                                                                     ['SRO Finance confidence'])
            try:
                ws.cell(row=row_num, column=13).value = master_one.data[project_name]['Last time at BICC']
                ws.cell(row=row_num, column=14).value = master_one.data[project_name]['Next at BICC']
            except KeyError:
                print('programme has crashed due to last time at BICC and Next and BICC keys not being inserted into'
                      'master')

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name in master_two.projects:
            ws.cell(row=row_num, column=5).value = convert_rag_text(master_two.data[project_name]['Departmental DCA'])

    # Highlight cells that contain RAG text, with background and text the same colour. column E.

    ag_text = Font(color="00a5b700")
    ag_fill = PatternFill(bgColor="00a5b700")
    dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    rule = Rule(type="containsText", operator="containsText", text="A/G", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A/G",e1)))']
    ws.conditional_formatting.add('e1:e100', rule)

    ar_text = Font(color="00f97b31")
    ar_fill = PatternFill(bgColor="00f97b31")
    dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    rule = Rule(type="containsText", operator="containsText", text="A/R", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A/R",e1)))']
    ws.conditional_formatting.add('e1:e100', rule)

    red_text = Font(color="00fc2525")
    red_fill = PatternFill(bgColor="00fc2525")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="R", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("R",E1)))']
    ws.conditional_formatting.add('E1:E100', rule)

    green_text = Font(color="0017960c")
    green_fill = PatternFill(bgColor="0017960c")
    dxf = DifferentialStyle(font=green_text, fill=green_fill)
    rule = Rule(type="containsText", operator="containsText", text="G", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("G",e1)))']
    ws.conditional_formatting.add('E1:E100', rule)

    amber_text = Font(color="00fce553")
    amber_fill = PatternFill(bgColor="00fce553")
    dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    rule = Rule(type="containsText", operator="containsText", text="A", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A",e1)))']
    ws.conditional_formatting.add('e1:e100', rule)

    # highlight cells in column g

    ag_text = Font(color="000000")
    ag_fill = PatternFill(bgColor="00a5b700")
    dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    rule = Rule(type="containsText", operator="containsText", text="A/G", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A/G",g1)))']
    ws.conditional_formatting.add('g1:g100', rule)

    ar_text = Font(color="000000")
    ar_fill = PatternFill(bgColor="00f97b31")
    dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    rule = Rule(type="containsText", operator="containsText", text="A/R", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A/R",g1)))']
    ws.conditional_formatting.add('g1:g100', rule)

    red_text = Font(color="000000")
    red_fill = PatternFill(bgColor="00fc2525")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="R", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("R",g1)))']
    ws.conditional_formatting.add('g1:g100', rule)

    green_text = Font(color="000000")
    green_fill = PatternFill(bgColor="0017960c")
    dxf = DifferentialStyle(font=green_text, fill=green_fill)
    rule = Rule(type="containsText", operator="containsText", text="G", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("G",g1)))']
    ws.conditional_formatting.add('g1:g100', rule)

    amber_text = Font(color="000000")
    amber_fill = PatternFill(bgColor="00fce553")
    dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    rule = Rule(type="containsText", operator="containsText", text="A", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A",g1)))']
    ws.conditional_formatting.add('g1:g100', rule)

    # highlight cells in column H

    ag_text = Font(color="000000")
    ag_fill = PatternFill(bgColor="00a5b700")
    dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    rule = Rule(type="containsText", operator="containsText", text="A/G", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A/G",h1)))']
    ws.conditional_formatting.add('h1:h100', rule)

    ar_text = Font(color="000000")
    ar_fill = PatternFill(bgColor="00f97b31")
    dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    rule = Rule(type="containsText", operator="containsText", text="A/R", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A/R",h1)))']
    ws.conditional_formatting.add('h1:h100', rule)

    red_text = Font(color="000000")
    red_fill = PatternFill(bgColor="00fc2525")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="R", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("R",h1)))']
    ws.conditional_formatting.add('h1:h100', rule)

    green_text = Font(color="000000")
    green_fill = PatternFill(bgColor="0017960c")
    dxf = DifferentialStyle(font=green_text, fill=green_fill)
    rule = Rule(type="containsText", operator="containsText", text="G", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("G",h1)))']
    ws.conditional_formatting.add('h1:h100', rule)

    amber_text = Font(color="000000")
    amber_fill = PatternFill(bgColor="00fce553")
    dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    rule = Rule(type="containsText", operator="containsText", text="A", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A",h1)))']
    ws.conditional_formatting.add('h1:h100', rule)

    # highlight cells in column H

    ag_text = Font(color="000000")
    ag_fill = PatternFill(bgColor="00a5b700")
    dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    rule = Rule(type="containsText", operator="containsText", text="A/G", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A/G",l1)))']
    ws.conditional_formatting.add('l1:l100', rule)

    ar_text = Font(color="000000")
    ar_fill = PatternFill(bgColor="00f97b31")
    dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    rule = Rule(type="containsText", operator="containsText", text="A/R", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A/R",l1)))']
    ws.conditional_formatting.add('l1:l100', rule)

    red_text = Font(color="000000")
    red_fill = PatternFill(bgColor="00fc2525")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="R", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("R",l1)))']
    ws.conditional_formatting.add('l1:l100', rule)

    green_text = Font(color="000000")
    green_fill = PatternFill(bgColor="0017960c")
    dxf = DifferentialStyle(font=green_text, fill=green_fill)
    rule = Rule(type="containsText", operator="containsText", text="G", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("G",l1)))']
    ws.conditional_formatting.add('l1:l100', rule)

    amber_text = Font(color="000000")
    amber_fill = PatternFill(bgColor="00fce553")
    dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    rule = Rule(type="containsText", operator="containsText", text="A", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("A",l1)))']
    ws.conditional_formatting.add('l1:l100', rule)


    # Highlight cells that contain RAG text, with background and black text columns G to L.
    # ag_text = Font(color="000000")
    # ag_fill = PatternFill(bgColor="00a5b700")
    # dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    # rule = Rule(type="uniqueValues", operator="equal", text="A/G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/G",G1)))']
    # ws.conditional_formatting.add('G1:L100', rule)
    #
    # ar_text = Font(color="000000")
    # ar_fill = PatternFill(bgColor="00f97b31")
    # dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    # rule = Rule(type="uniqueValues", operator="equal", text="A/R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/R",G1)))']
    # ws.conditional_formatting.add('G1:L100', rule)
    #
    # red_text = Font(color="000000")
    # red_fill = PatternFill(bgColor="00fc2525")
    # dxf = DifferentialStyle(font=red_text, fill=red_fill)
    # rule = Rule(type="uniqueValues", operator="equal", text="R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("R",G1)))']
    # ws.conditional_formatting.add('G1:L100', rule)
    #
    # green_text = Font(color="000000")
    # green_fill = PatternFill(bgColor="0017960c")
    # dxf = DifferentialStyle(font=green_text, fill=green_fill)
    # rule = Rule(type="uniqueValues", operator="equal", text="G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("Green",G1)))']
    # ws.conditional_formatting.add('G1:L100', rule)
    #
    # amber_text = Font(color="000000")
    # amber_fill = PatternFill(bgColor="00fce553")
    # dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    # rule = Rule(type="uniqueValues", operator="equal", text="A", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A",G1)))']
    # ws.conditional_formatting.add('G1:L100', rule)

    # highlighting new projects
    red_text = Font(color="00fc2525")
    white_fill = PatternFill(bgColor="000000")
    dxf = DifferentialStyle(font=red_text, fill=white_fill)
    rule = Rule(type="uniqueValues", operator="equal", text="NEW", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("NEW",F1)))']
    ws.conditional_formatting.add('F1:F100', rule)

    # assign the icon set to a rule
    first = FormatObject(type='num', val=-1)
    second = FormatObject(type='num', val=0)
    third = FormatObject(type='num', val=1)
    iconset = IconSet(iconSet='3Arrows', cfvo=[first, second, third], percent=None, reverse=None)
    rule = Rule(type='iconSet', iconSet=iconset)
    ws.conditional_formatting.add('F1:F100', rule)

    # change text in last at next at BICC column
    for row_num in range(2, ws.max_row + 1):
        if ws.cell(row=row_num, column=13).value == '-2 weeks':
            ws.cell(row=row_num, column=13).value = 'Last BICC'
        if ws.cell(row=row_num, column=13).value == '2 weeks':
            ws.cell(row=row_num, column=13).value = 'Next BICC'
        if ws.cell(row=row_num, column=13).value == 'Today':
            ws.cell(row=row_num, column=13).value = 'This BICC'
        if ws.cell(row=row_num, column=14).value == '-2 weeks':
            ws.cell(row=row_num, column=14).value = 'Last BICC'
        if ws.cell(row=row_num, column=14).value == '2 weeks':
            ws.cell(row=row_num, column=14).value = 'Next BICC'
        if ws.cell(row=row_num, column=14).value == 'Today':
            ws.cell(row=row_num, column=14).value = 'This BICC'

            # highlight text in bold
    ft = Font(bold=True)
    for row_num in range(2, ws.max_row + 1):
        lis = ['This week', 'Next week', 'Last week', 'Two weeks',
               'Two weeks ago', 'This mth', 'Last mth', 'Next mth',
               '2 mths', '3 mths', '-2 mths', '-3 mths', '-2 weeks',
               'Today', 'Last BICC', 'Next BICC', 'This BICC',
               'Later this mth']
        if ws.cell(row=row_num, column=10).value in lis:
            ws.cell(row=row_num, column=10).font = ft
        if ws.cell(row=row_num, column=11).value in lis:
            ws.cell(row=row_num, column=11).font = ft
        if ws.cell(row=row_num, column=13).value in lis:
            ws.cell(row=row_num, column=13).font = ft
        if ws.cell(row=row_num, column=14).value in lis:
            ws.cell(row=row_num, column=14).font = ft
    return wb


''' RUNNING THE PROGRAMME '''

'''ONE. Provide file path to dashboard master'''
dashboard_master = load_workbook('C:\\Users\\Standalone\\general\\masters folder\\portfolio_dashboards\\master.xlsx')

'''TWO. Provide list of projects on which to provide analysis'''
project_list = q2_1920.projects

'''THREE. Specify data of BICC that is discussing the report. NOTE: Python date format is (YYYY,MM,DD)'''
bicc_date = datetime.datetime(2019, 9, 9)

'''FOUR. place arguments into the place_in_excle function and provide file path for saving output wb'''
dashboard_completed = place_in_excel(q2_1920, q1_1920, dashboard_master)
dashboard_completed.save('C:\\Users\\Standalone\\general\\masters folder\\portfolio_dashboards\\test.xlsx')