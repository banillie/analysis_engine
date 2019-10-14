'''
Programme produces outputs (figures) for building the speed dials page

outputs:
programme prints out into MS word high level numbers for speed dials
overall SRO DCA changes between quarters - in MS word. Typically current and last. But can be adapted to whichever two
quarters are of interest
overall SRO finance confidence changes between quarters - same as for overall sro dca
overall SRO benefits confidence changes between quarters - same as for overall sro dca
overall SRO resource confidence changes between quarters - same as for overall sro dca

follow instructions below.

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.

TODO: have calculate overall dials information included as part of the ms word output documents, not just a print out in the console.
TODO: better solution than having to specifying file paths in the run_programme function.
'''

import docx
from analysis.data import q2_1920, q1_1920

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

    print(output)
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
        #print(project_name, dca_change_master[project_name])
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

def calculate_overall_dials(dca_change_master):
    empty_list = []
    for x in dca_change_master:
        empty_list.append(dca_change_master[x]['DCA'])

    '''this section is necessary as some dcas have 5 options, while others have three'''
    if len(list(set(empty_list))) > 3:
        count_list = []
        red = empty_list.count('Red')
        count_list.append(('Red', red))
        amber_red = empty_list.count('Amber/Red')
        count_list.append(('Amber/Red', amber_red))
        amber = empty_list.count('Amber')
        count_list.append(('Amber', amber))
        amber_green = empty_list.count('Amber/Green')
        count_list.append(('Amber/Green', amber_green))
        green = empty_list.count('Green')
        count_list.append(('Green', green))

        print(count_list)

        total = 0
        for i in range(0, len(count_list)):
            total += (count_list[i][1])

        print('total number of projects ' + str(total))

        a = count_list[0][1] * 0
        b = count_list[1][1] * 25
        c = count_list[2][1] * 50
        d = count_list[3][1] * 75
        e = count_list[4][1] * 100

        score = a + b + c + d + e
        maximum = total * 100

        result = score / maximum

        print(result)



    if len(list(set(empty_list))) <= 3:
        count_list = []
        red = empty_list.count('Red')
        count_list.append(('Red', red))
        amber = empty_list.count('Amber')
        count_list.append(('Amber', amber))
        green = empty_list.count('Green')
        count_list.append(('Green', green))

        print(count_list)

        total = 0
        for i in range(0, len(count_list)):
            total += (count_list[i][1])

        print('total number of projects ' + str(total))

        a = count_list[0][1] * 0
        b = count_list[1][1] * 50
        c = count_list[2][1] * 100

        score = a + b + c
        maximum = total * 100

        result = score / maximum

        print(result)

def run_programme(master_one, master_two):
    '''
    Function that runs the programme.

    :param master_one: master quarter of interest one. typically latest quarter
    :param master_two: master quarter of interest two. typically last quarter
    :return: four ms word documents.
    '''

    sro_dca = calculate_dca_change('Departmental DCA', master_one, master_two)
    finance_dca = calculate_dca_change('SRO Finance confidence', master_one, master_two)
    resource_dca = calculate_dca_change('Overall Resource DCA - Now', master_one, master_two)
    benefits_dca = calculate_dca_change('SRO Benefits RAG', master_one, master_two)
    schedule_dca = calculate_overall_dials('SRO Schedule Confidence', master_one, master_two)

    print('DCA')
    calculate_overall_dials(sro_dca)
    print('Finance')
    calculate_overall_dials(finance_dca)
    print('Resource')
    calculate_overall_dials(resource_dca)
    print('Benefits')
    calculate_overall_dials(benefits_dca)
    print('Schedule')
    calculate_overall_dials(schedule_dca)

    overall = print_dca_change(sro_dca)
    finance = print_dca_change(finance_dca)
    resource = print_dca_change(resource_dca)
    benefits = print_dca_change(benefits_dca)
    schedule = print_dca_change(schedule_dca)

    overall.save('C:\\Users\\Standalone\\general\\masters folder\\speed_dials\\q2_1920_overall_dca.docx')
    finance.save('C:\\Users\\Standalone\\general\\masters folder\\speed_dials\\q2_1920_finance_dca.docx')
    resource.save('C:\\Users\\Standalone\\general\\masters folder\\speed_dials\\q2_1920_resource_dca.docx')
    benefits.save('C:\\Users\\Standalone\\general\\masters folder\\speed_dials\\q2_1920_benefits_dca.docx')
    schedule.save('C:\\Users\\Standalone\\general\\masters folder\\speed_dials\\q2_1920_schedule_dca.docx')


''' RUNNING PROGRAMME '''

''' Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.'''

'''ONE. the below programme simply need to be run. now commands required. However, need to find a way to easily change
file paths

Note. file paths for saving outputs currently need to be specified in the programme itself'''

run_programme(q2_1920, q1_1920)