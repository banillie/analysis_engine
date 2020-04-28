'''
Programme produces outputs (figures) for building the speed dials page

outputs:
- Programme saves (into output dir) Word documents project changes for each dial.
- It prints out in the calculations for overall dial position into excel output titled dials.

Follow instructions at end of code.

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.
'''

import docx
from analysis.data import list_of_masters_all, root_path
from collections import Counter
from openpyxl import Workbook

def calculate_dca_change(category, master_one, master_two):
    '''
    function that calculates whether dca ratings have changed and by how much
    :param category: dca category e.g. delivery, finance
    :param master_one: master quarter of interest one. typically latest quarter
    :param master_two: master quarter of interest two. typically last quarter
    :return:
    '''
    output = {}
    for project_name in master_one.projects:
        output[project_name] = {}
        holding_dict_1 = master_one.data[project_name]
        a = holding_dict_1[category]
        output[project_name]['DCA'] = a
        try:
            holding_dict_2 = master_two.data[project_name]
            b = holding_dict_2[category]
            output[project_name]['DCA_lq'] = b
        except KeyError:
            b = 'not reporting'
            output[project_name]['DCA_lq'] = b
        if a == b:
            output[project_name]['Change'] = 0
        elif a != b:
            if output[project_name]['DCA_lq'] == 'not reporting':
                output[project_name]['Change'] = 0
            elif output[project_name]['DCA'] == None:  # quick hack fix here for handling projects that remove dca reporting
                output[project_name]['Change'] = 0
            elif output[project_name]['DCA'] == 'Green':
                if output[project_name]['DCA_lq'] == 'Amber/Green':
                    output[project_name]['Change'] = 1
                if output[project_name]['DCA_lq'] == 'Amber':
                    output[project_name]['Change'] = 2
                if output[project_name]['DCA_lq'] == 'Amber/Red':
                    output[project_name]['Change'] = 3
                if output[project_name]['DCA_lq'] == 'Red':
                    output[project_name]['Change'] = 4
                if output[project_name]['DCA_lq'] == None:
                    output[project_name]['Change'] = 5
            elif output[project_name]['DCA'] == 'Amber/Green':
                if output[project_name]['DCA_lq'] == 'Green':
                    output[project_name]['Change'] = -1
                if output[project_name]['DCA_lq'] == 'Amber':
                    output[project_name]['Change'] = 1
                if output[project_name]['DCA_lq'] == 'Amber/Red':
                    output[project_name]['Change'] = 2
                if output[project_name]['DCA_lq'] == 'Red':
                    output[project_name]['Change'] = 3
                if output[project_name]['DCA_lq'] == None:
                    output[project_name]['Change'] = 5
            elif output[project_name]['DCA'] == 'Amber':
                if output[project_name]['DCA_lq'] == 'Green':
                    output[project_name]['Change'] = -2
                if output[project_name]['DCA_lq'] == 'Amber/Green':
                    output[project_name]['Change'] = -1
                if output[project_name]['DCA_lq'] == 'Amber/Red':
                    output[project_name]['Change'] = 1
                if output[project_name]['DCA_lq'] == 'Red':
                    output[project_name]['Change'] = 2
                if output[project_name]['DCA_lq'] == None:
                    output[project_name]['Change'] = 5
            elif output[project_name]['DCA'] == 'Amber/Red':
                if output[project_name]['DCA_lq'] == 'Green':
                    output[project_name]['Change'] = -3
                if output[project_name]['DCA_lq'] == 'Amber/Green':
                    output[project_name]['Change'] = -2
                if output[project_name]['DCA_lq'] == 'Amber':
                    output[project_name]['Change'] = -1
                if output[project_name]['DCA_lq'] == 'Red':
                    output[project_name]['Change'] = 1
                if output[project_name]['DCA_lq'] == None:
                    output[project_name]['Change'] = 5
            elif output[project_name]['DCA'] == 'Red':
                if output[project_name]['DCA_lq'] == 'Green':
                    output[project_name]['Change'] = -4
                if output[project_name]['DCA_lq'] == 'Amber/Green':
                    output[project_name]['Change'] = -3
                if output[project_name]['DCA_lq'] == 'Amber':
                    output[project_name]['Change'] = -2
                if output[project_name]['DCA_lq'] == 'Amber/Red':
                    output[project_name]['Change'] = -1
                if output[project_name]['DCA_lq'] == None:
                    output[project_name]['Change'] = 5

    #print(output)
    return output

def print_dca_change(dca_change_master):
    doc = docx.Document()
    title = 'Confidence changes this quarter'
    top = doc.add_paragraph()
    top.add_run(title).bold = True
    doc.add_paragraph()
    sub_head = 'Decrease (in order of size of change)'
    sub = doc.add_paragraph()
    sub.add_run(sub_head).bold = True
    down = 0
    for project_name in dca_change_master:
        # print(project_name)
        # print(project_name, dca_change_master[project_name])
        if dca_change_master[project_name]['Change'] == -4:
            p = doc.add_paragraph()
            a = dca_change_master[project_name]['DCA']
            b = dca_change_master[project_name]['DCA_lq']
            c = str(project_name)
            down += 1
            d = str(down) + '. ' + c
            e = ': change from ' + str(b) + ' to ' + str(a)
            p.add_run(d).bold = True
            p.add_run(e)
    for project_name in dca_change_master:
        if dca_change_master[project_name]['Change'] == -3:
            p = doc.add_paragraph()
            a = dca_change_master[project_name]['DCA']
            b = dca_change_master[project_name]['DCA_lq']
            c = str(project_name)
            down += 1
            d = str(down) + '. ' + c
            e = ': change from ' + str(b) + ' to ' + str(a)
            p.add_run(d).bold = True
            p.add_run(e)
    for project_name in dca_change_master:
        if dca_change_master[project_name]['Change'] == -2:
            p = doc.add_paragraph()
            a = dca_change_master[project_name]['DCA']
            b = dca_change_master[project_name]['DCA_lq']
            c = str(project_name)
            down += 1
            d = str(down) + '. ' + c
            e = ': change from ' + str(b) + ' to ' + str(a)
            p.add_run(d).bold = True
            p.add_run(e)
    for project_name in dca_change_master:
        if dca_change_master[project_name]['Change'] == -1:
            p = doc.add_paragraph()
            a = dca_change_master[project_name]['DCA']
            b = dca_change_master[project_name]['DCA_lq']
            c = str(project_name)
            down += 1
            d = str(down) + '. ' + c
            e = ': change from ' + str(b) + ' to ' + str(a)
            p.add_run(d).bold = True
            p.add_run(e)
    doc.add_paragraph()
    g = doc.add_paragraph()
    totals_line = str(down) + ' project(s) have decreased in total'
    g.add_run(totals_line).bold = True
    doc.add_paragraph()
    sub_head_2 = 'Increase (in order of size of change)'
    sub_2 = doc.add_paragraph()
    sub_2.add_run(sub_head_2).bold = True
    up = 0
    for project_name in dca_change_master:
        if dca_change_master[project_name]['Change'] == 4:
            p = doc.add_paragraph()
            a = dca_change_master[project_name]['DCA']
            b = dca_change_master[project_name]['DCA_lq']
            c = str(project_name)
            up += 1
            d = str(up) + '. ' + c
            e = ': change from ' + str(b) + ' to ' + str(a)
            p.add_run(d).bold = True
            p.add_run(e)
    for project_name in dca_change_master:
        if dca_change_master[project_name]['Change'] == 3:
            p = doc.add_paragraph()
            a = dca_change_master[project_name]['DCA']
            b = dca_change_master[project_name]['DCA_lq']
            c = str(project_name)
            up += 1
            d = str(up) + '. ' + c
            e = ': change from ' + str(b) + ' to ' + str(a)
            p.add_run(d).bold = True
            p.add_run(e)
    for project_name in dca_change_master:
        if dca_change_master[project_name]['Change'] == 2:
            p = doc.add_paragraph()
            a = dca_change_master[project_name]['DCA']
            b = dca_change_master[project_name]['DCA_lq']
            up += 1
            c = str(project_name)
            d = str(up) + '. ' + c
            e = ': change from ' + str(b) + ' to ' + str(a)
            p.add_run(d).bold = True
            p.add_run(e)
    for project_name in dca_change_master:
        if dca_change_master[project_name]['Change'] == 1:
            p = doc.add_paragraph()
            a = dca_change_master[project_name]['DCA']
            b = dca_change_master[project_name]['DCA_lq']
            c = str(project_name)
            up += 1
            d = str(up) + '. ' + c
            e = ': change from ' + str(b) + ' to ' + str(a)
            p.add_run(d).bold = True
            p.add_run(e)
    doc.add_paragraph()
    totals_line_2 = str(up) + ' project(s) have increased in total'
    h = doc.add_paragraph()
    h.add_run(totals_line_2).bold = True

    return doc

def sort_by_rag(quarter_data, dca_type):

    rag_list = []
    for project_name in list_of_masters_all[0].projects:
        rag = quarter_data.data[project_name][dca_type]
        if rag is not None:
            rag_list.append((project_name, rag))
        else:
            pass

    rag_list_sorted = sorted(rag_list, key=lambda x:x[1])

    return rag_list_sorted

def calculate_overall_dials():

    wb = Workbook()
    ws = wb.active

    l_data = list_of_masters_all[0]

    dca_rags = Counter(x[1] for x in sort_by_rag(l_data, 'Departmental DCA'))
    fin_rags = Counter(x[1] for x in sort_by_rag(l_data, 'SRO Finance confidence'))
    ben_rags = Counter(x[1] for x in sort_by_rag(l_data, 'SRO Benefits RAG'))
    schedule_rags = Counter(x[1] for x in sort_by_rag(l_data, 'SRO Schedule Confidence'))
    resource_rags = Counter(x[1] for x in sort_by_rag(l_data, 'Overall Resource DCA - Now'))

    rag_list = [dca_rags, fin_rags, ben_rags, schedule_rags, resource_rags]
    rag_cl_list = ['Red', 'Amber/Red', 'Amber', 'Amber/Green', 'Green']
    confidence_type_list = ['Overall', 'Finance', 'Benefits', 'Schedule', 'Resources']

    for x, colour in enumerate(rag_cl_list):
        for i, rag_dict in enumerate(rag_list):
            ws.cell(row=x+2, column=i+2).value = rag_dict[colour]

    for x, rag_dict in enumerate(rag_list):
        ws.cell(row=7, column=x+2).value = sum(rag_dict.values())

    for x, value in enumerate(rag_cl_list):
        ws.cell(row=x+2, column=1).value = value

    for x, value in enumerate(confidence_type_list):
        ws.cell(row=1, column=x+2).value = value

    ws.cell(row=7, column=1).value = 'Total'
    ws.cell(row=1, column=1).value = 'DCA'

    return wb

def run_programme(quarter_output):

    sro_dca = calculate_dca_change('Departmental DCA', list_of_masters_all[0], list_of_masters_all[1])
    finance_dca = calculate_dca_change('SRO Finance confidence', list_of_masters_all[0], list_of_masters_all[1])
    resource_dca = calculate_dca_change('Overall Resource DCA - Now', list_of_masters_all[0], list_of_masters_all[1])
    benefits_dca = calculate_dca_change('SRO Benefits RAG', list_of_masters_all[0], list_of_masters_all[1])
    schedule_dca = calculate_dca_change('SRO Schedule Confidence', list_of_masters_all[0], list_of_masters_all[1])

    output = calculate_overall_dials()
    output.save(root_path/'output/speed_dials_{}.xlsx'.format(quarter_output))

    overall = print_dca_change(sro_dca)
    finance = print_dca_change(finance_dca)
    resource = print_dca_change(resource_dca)
    benefits = print_dca_change(benefits_dca)
    schedule = print_dca_change(schedule_dca)

    overall.save(root_path/'output/overall_speed_dials_{}.docx'.format(quarter_output))
    finance.save(root_path/'output/financial_speed_dials_{}.docx'.format(quarter_output))
    resource.save(root_path/'output/resource_speed_dials_{}.docx'.format(quarter_output))
    benefits.save(root_path/'output/benefits_speed_dials_{}.docx'.format(quarter_output))
    schedule.save(root_path/'output/schedule_speed_dials_{}.docx'.format(quarter_output))


''' RUNNING PROGRAMME '''

'''enter into the run_programme function the quarter details for the output files e.g. _q4_1920 (note put underscore at 
front)'''
run_programme('q4_1920')