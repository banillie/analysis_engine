'''

In development... function for creating different dashboard in line with PDIP work

Input documents:
1) Dashboard master document - this is an excel file.
2) Master data for two quarters - this will usually be latest and previous quarter

output document:
3) Dashboard with all project data placed into dashboard and formatted correctly.

Instructions:
1) provide path to dashboard master
2) provide path to master data sets
3) change bicc_date variable
4) provide path and specify file name for output document
'''

from openpyxl import load_workbook
import datetime
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, IconSet, FormatObject
from engine_functions import all_milestone_data_bulk, concatenate_dates, up_or_down
from data import q1_1920, q4_1819

'''Function that creates dictionary with keys of interest'''
def inital_dict(project_name, master, key_list):
    upper_dictionaryionary = {}
    for name in project_name:
        lower_dictionary = {}

        try:
            p_data = master.data[name]

            for value in key_list:
                if value in p_data.keys():
                    lower_dictionary[value] = p_data[value]

        except KeyError:
            pass

        upper_dictionaryionary[name] = lower_dictionary

    return upper_dictionaryionary

def add_sop_pend_data(m_data, dict):

    for name in dict.keys():
        try:
            dict[name]['Start of Operation'] = m_data[name]['Start of Operation']
        except KeyError:
            print(name + ' no sop date')
            dict[name]['Start of Operation'] = None
        try:
            dict[name]['Project End Date'] = m_data[name]['Project End Date']
        except KeyError:
            print(name + ' no proj end date')
            dict[name]['Project End Date'] = None

    return dict

'''function for adding concatenated word strings to dictionary.
note probably don't need the above function now, but can tidy up later'''
def final_dict(dict_one, dict_two, con_list, dca_key):
    upper_dictionary = {}

    for name in dict_one:
        lower_dict = {}
        p_dict_one = dict_one[name]
        for key in p_dict_one:
            if key in con_list:
                try:
                    lower_dict[key] = concatenate_dates(p_dict_one[key])
                except TypeError:
                    try:
                        lower_dict[key] = concatenate_dates(tuple(p_dict_one[key])[0])
                    except TypeError:
                        lower_dict[key] = 'check data'
            else:
                lower_dict[key] = p_dict_one[key]

        try:
            lower_dict['Change'] = up_or_down(p_dict_one[dca_key], dict_two[name][dca_key])
        except KeyError:
            lower_dict['Change'] = 'NEW'

        upper_dictionary[name] = lower_dict

    return upper_dictionary

def convert_rag_text(dca_rating):

    if dca_rating == 'Green':
        return 'G'
    elif dca_rating == 'Amber/Green':
        return 'A/G'
    elif dca_rating == 'Amber':
        return 'A'
    elif dca_rating == 'Amber/Red':
        return 'A/R'
    elif dca_rating == 'Red':
        return 'R'

def convert_bc_stage_text(bc_stage):

    if bc_stage == 'Strategic Outline Case':
        return 'SOBC'
    elif bc_stage == 'Outline Business Case':
        return 'OBC'
    elif bc_stage == 'Full Business Case':
        return 'FBC'
    elif bc_stage == 'pre-Strategic Outline Case':
        return 'pre-SOBC'
    else:
        return bc_stage

def calculating_schedule_progression(proj_name, m_data):

    '''function that calculations the propotion of schedule that has been completed'''

    start_date = tuple(m_data[proj_name]['Start of Project'])[0]

    end_date = tuple(m_data[proj_name]['Project End Date'])[0]

    now = datetime.date.today()

    try:
        proj_length = (end_date - start_date).days
        remain_time = (end_date - now).days
        completed = round(100 - ((remain_time / proj_length) * 100), 0)
    except TypeError:
        completed = 'missing data'

    return completed

def calculating_cost_progression(master_data, proj_name):

    total = master_data.data[proj_name]['Total Forecast']


    pre_rdel = master_data.data[proj_name]['Pre 19-20 RDEL Forecast Total']
    if pre_rdel == None:
        pre_rdel = 0
    pre_cdel = master_data.data[proj_name]['Pre 19-20 CDEL Forecast Total']
    if pre_cdel == None:
        pre_cdel = 0
    pre_ngov = master_data.data[proj_name]['Pre 19-20 Forecast Non-Gov']
    if pre_ngov == None:
        pre_ngov = 0
    pre_spend = pre_rdel + pre_cdel + pre_ngov

    try:
        output = round((pre_spend / total) * 100, 0)
    except ZeroDivisionError:
        output = 0

    return output



'''function that places all information into the summary dashboard sheet'''
def placing_excel(dict_one, dict_two, altered_dict, milestone_dict):

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        print(project_name)
        if project_name in dict_one.projects:
            ws.cell(row=row_num, column=4).value = round(dict_one.data[project_name]['Total Forecast'], 0)
            ws.cell(row=row_num, column=5).value = dict_one.data[project_name]['Adjusted Benefits Cost Ratio (BCR)']
            ws.cell(row=row_num, column=6).value = tuple(altered_dict[project_name]['Project End Date'])[0]
            ws.cell(row=row_num, column=7).value = convert_rag_text(dict_one.data[project_name]['Departmental DCA'])
            ws.cell(row=row_num, column=8).value = altered_dict[project_name]['Change']
            ws.cell(row=row_num, column=10).value = calculating_schedule_progression(project_name, milestone_dict)
            ws.cell(row=row_num, column=11).value = calculating_cost_progression(dict_one, project_name)
            ws.cell(row=row_num, column=12).value = \
                dict_one.data[project_name]['Optimism Bias Percentage Used in Cost Baselines']
            ws.cell(row=row_num, column=13).value = dict_one.data[project_name]['Overall figure for Optimism Bias (£m)']
            ws.cell(row=row_num, column=14).value = \
                dict_one.data[project_name]['Built in contingency (% of Whole Life Cost)']
            ws.cell(row=row_num, column=15).value = dict_one.data[project_name]['Overall contingency (£m)']

            # ws.cell(row=row_num, column=8).value = convert_rag_text(dict_one[project_name]['GMPP - IPA DCA last quarter'])
            # ws.cell(row=row_num, column=9).value = convert_bc_stage_text(dict_one[project_name]['BICC approval point'])
            # ws.cell(row=row_num, column=10).value = dict_one[project_name]['Start of Operation']
            # ws.cell(row=row_num, column=11).value = dict_one[project_name]['Project End Date']
            # ws.cell(row=row_num, column=12).value = convert_rag_text(dict_one[project_name]['SRO Finance confidence'])
            # ws.cell(row=row_num, column=13).value = dict_one[project_name]['Last time at BICC']
            # ws.cell(row=row_num, column=14).value = dict_one[project_name]['Next at BICC']

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name in dict_two.projects:
            ws.cell(row=row_num, column=9).value = convert_rag_text(dict_two.data[project_name]['Departmental DCA'])

    # Highlight cells that contain RAG text, with background and text the same colour. column G.

    # rag_txt_list = ["A/G", "A/R", "R", "G", "A"]
    #
    # for name in rag_txt_list:
    #     ag_text = Font(color="00a5b700")
    #     ag_fill = PatternFill(bgColor="00a5b700")
    #     dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    #     rule = Rule(type="containsText", operator="containsText", text=name, dxf=dxf)
    #     rule.formula = ['NOT(ISERROR(SEARCH('+name+',g1)))']
    #     ws.conditional_formatting.add('g1:g100', rule)
    #
    # for name in rag_txt_list:
    #     ag_text = Font(color="00a5b700")
    #     ag_fill = PatternFill(bgColor="00a5b700")
    #     dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    #     rule = Rule(type="containsText", operator="containsText", text=name, dxf=dxf)
    #     rule.formula = ['NOT(ISERROR(SEARCH('+name+',i1)))']
    #     ws.conditional_formatting.add('i1:i100', rule)

    # ar_text = Font(color="00f97b31")
    # ar_fill = PatternFill(bgColor="00f97b31")
    # dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A/R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/R",g1)))']
    # ws.conditional_formatting.add('g1:g100', rule)
    #
    # red_text = Font(color="00fc2525")
    # red_fill = PatternFill(bgColor="00fc2525")
    # dxf = DifferentialStyle(font=red_text, fill=red_fill)
    # rule = Rule(type="containsText", operator="containsText", text="R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("R",E1)))']
    # ws.conditional_formatting.add('E1:E100', rule)
    #
    # green_text = Font(color="0017960c")
    # green_fill = PatternFill(bgColor="0017960c")
    # dxf = DifferentialStyle(font=green_text, fill=green_fill)
    # rule = Rule(type="containsText", operator="containsText", text="G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("G",e1)))']
    # ws.conditional_formatting.add('E1:E100', rule)
    #
    # amber_text = Font(color="00fce553")
    # amber_fill = PatternFill(bgColor="00fce553")
    # dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A",e1)))']
    # ws.conditional_formatting.add('e1:e100', rule)

    # highlight cells in column g

    # ag_text = Font(color="000000")
    # ag_fill = PatternFill(bgColor="00a5b700")
    # dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A/G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/G",g1)))']
    # ws.conditional_formatting.add('g1:g100', rule)
    #
    # ar_text = Font(color="000000")
    # ar_fill = PatternFill(bgColor="00f97b31")
    # dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A/R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/R",g1)))']
    # ws.conditional_formatting.add('g1:g100', rule)
    #
    # red_text = Font(color="000000")
    # red_fill = PatternFill(bgColor="00fc2525")
    # dxf = DifferentialStyle(font=red_text, fill=red_fill)
    # rule = Rule(type="containsText", operator="containsText", text="R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("R",g1)))']
    # ws.conditional_formatting.add('g1:g100', rule)
    #
    # green_text = Font(color="000000")
    # green_fill = PatternFill(bgColor="0017960c")
    # dxf = DifferentialStyle(font=green_text, fill=green_fill)
    # rule = Rule(type="containsText", operator="containsText", text="G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("G",g1)))']
    # ws.conditional_formatting.add('g1:g100', rule)
    #
    # amber_text = Font(color="000000")
    # amber_fill = PatternFill(bgColor="00fce553")
    # dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A",g1)))']
    # ws.conditional_formatting.add('g1:g100', rule)
    #
    # # highlight cells in column H
    #
    # ag_text = Font(color="000000")
    # ag_fill = PatternFill(bgColor="00a5b700")
    # dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A/G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/G",h1)))']
    # ws.conditional_formatting.add('h1:h100', rule)
    #
    # ar_text = Font(color="000000")
    # ar_fill = PatternFill(bgColor="00f97b31")
    # dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A/R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/R",h1)))']
    # ws.conditional_formatting.add('h1:h100', rule)
    #
    # red_text = Font(color="000000")
    # red_fill = PatternFill(bgColor="00fc2525")
    # dxf = DifferentialStyle(font=red_text, fill=red_fill)
    # rule = Rule(type="containsText", operator="containsText", text="R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("R",h1)))']
    # ws.conditional_formatting.add('h1:h100', rule)
    #
    # green_text = Font(color="000000")
    # green_fill = PatternFill(bgColor="0017960c")
    # dxf = DifferentialStyle(font=green_text, fill=green_fill)
    # rule = Rule(type="containsText", operator="containsText", text="G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("G",h1)))']
    # ws.conditional_formatting.add('h1:h100', rule)
    #
    # amber_text = Font(color="000000")
    # amber_fill = PatternFill(bgColor="00fce553")
    # dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A",h1)))']
    # ws.conditional_formatting.add('h1:h100', rule)
    #
    # # highlight cells in column H
    #
    # ag_text = Font(color="000000")
    # ag_fill = PatternFill(bgColor="00a5b700")
    # dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A/G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/G",l1)))']
    # ws.conditional_formatting.add('l1:l100', rule)
    #
    # ar_text = Font(color="000000")
    # ar_fill = PatternFill(bgColor="00f97b31")
    # dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A/R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A/R",l1)))']
    # ws.conditional_formatting.add('l1:l100', rule)
    #
    # red_text = Font(color="000000")
    # red_fill = PatternFill(bgColor="00fc2525")
    # dxf = DifferentialStyle(font=red_text, fill=red_fill)
    # rule = Rule(type="containsText", operator="containsText", text="R", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("R",l1)))']
    # ws.conditional_formatting.add('l1:l100', rule)
    #
    # green_text = Font(color="000000")
    # green_fill = PatternFill(bgColor="0017960c")
    # dxf = DifferentialStyle(font=green_text, fill=green_fill)
    # rule = Rule(type="containsText", operator="containsText", text="G", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("G",l1)))']
    # ws.conditional_formatting.add('l1:l100', rule)
    #
    # amber_text = Font(color="000000")
    # amber_fill = PatternFill(bgColor="00fce553")
    # dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    # rule = Rule(type="containsText", operator="containsText", text="A", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("A",l1)))']
    # ws.conditional_formatting.add('l1:l100', rule)


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
    # red_text = Font(color="00fc2525")
    # white_fill = PatternFill(bgColor="000000")
    # dxf = DifferentialStyle(font=red_text, fill=white_fill)
    # rule = Rule(type="uniqueValues", operator="equal", text="NEW", dxf=dxf)
    # rule.formula = ['NOT(ISERROR(SEARCH("NEW",F1)))']
    # ws.conditional_formatting.add('F1:F100', rule)

    # assign the icon set to a rule
    first = FormatObject(type='num', val=-1)
    second = FormatObject(type='num', val=0)
    third = FormatObject(type='num', val=1)
    iconset = IconSet(iconSet='3Arrows', cfvo=[first, second, third], percent=None, reverse=None)
    rule = Rule(type='iconSet', iconSet=iconset)
    ws.conditional_formatting.add('H1:H100', rule)

    return wb


'''keys of interest for current quarter'''
dash_keys = ['Total Forecast', 'Departmental DCA', 'BICC approval point',
            'Project Lifecycle Stage', 'SRO Finance confidence', 'Last time at BICC', 'Next at BICC',
             'GMPP - IPA DCA last quarter']

'''key of interest for previous quarter'''
dash_keys_previous_quarter = ['Departmental DCA']

keys_to_concatenate = ['Start of Operation', 'Last time at BICC',
                       'Next at BICC']

'''1) Provide file to empty dashboard document'''
wb = load_workbook(
    'C:\\Users\\Standalone\\general\\masters folder\\portfolio_dashboards\\pdip_dashboard_poc_master.xlsx')
ws = wb.active

p_names = q1_1920.projects

'''3) Specify data of bicc that is discussing the report. NOTE: Python date format is (YYYY,MM,DD)'''
bicc_date = datetime.datetime(2019, 9, 9)


latest_q_dict = inital_dict(p_names, q1_1920, dash_keys)
last_q_dict = inital_dict(p_names, q4_1819, dash_keys_previous_quarter)
m_data = all_milestone_data_bulk(p_names, q1_1920)
latest_q_dict_2 = add_sop_pend_data(m_data, latest_q_dict)
merged_dict = final_dict(latest_q_dict_2, last_q_dict, keys_to_concatenate, 'Departmental DCA')

wb = placing_excel(q1_1920, q4_1819, merged_dict, m_data)

'''4) provide file path and specific name of output file.'''
wb.save('C:\\Users\\Standalone\\general\\masters folder\\portfolio_dashboards\\pdip_dashboard_poc_completed.xlsx')