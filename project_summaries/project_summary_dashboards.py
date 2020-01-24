'''
Programme that compiles project dashboards/summary sheets.

Input:
1) four quarters worth of data

Output:
1) MS word document in structure of summary sheet / dashboard - with some areas missing, see below.

Supplementary programmes that need to be run to build charts for summary pages. Charts should be built and cut and paste
into dashboards/summary sheets:
1) project_financial_profile.py . For financial charts
2) milestone_comparison_3_qrts_proj.py . For milestones tables

'''

from docx import Document
import datetime
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import Cm, RGBColor
import difflib

from analysis.engine_functions import convert_rag_text, all_milestone_data_bulk
from analysis.data import latest_quarter_project_names, list_of_masters_all, root_path


def cell_colouring(cell, colour):
    '''
    Function that handles cell colouring
    cell: cell reference
    color: colour reference
    '''

    try:
        if colour == 'R':
            colour = parse_xml(r'<w:shd {} w:fill="cb1f00"/>'.format(nsdecls('w')))
        elif colour == 'A/R':
            colour = parse_xml(r'<w:shd {} w:fill="f97b31"/>'.format(nsdecls('w')))
        elif colour == 'A':
            colour = parse_xml(r'<w:shd {} w:fill="fce553"/>'.format(nsdecls('w')))
        elif colour == 'A/G':
            colour = parse_xml(r'<w:shd {} w:fill="a5b700"/>'.format(nsdecls('w')))
        elif colour == 'G':
            colour = parse_xml(r'<w:shd {} w:fill="17960c"/>'.format(nsdecls('w')))

        cell._tc.get_or_add_tcPr().append(colour)

    except TypeError:
        pass

def compare_text_showall(text_1, text_2, doc):
    '''
    Function places text into doc highlighting all changes.
    text_1: latest text. string.
    text_2: last text. string
    doc: word doc
    '''

    comp = difflib.Differ()
    diff = list(comp.compare(text_2.split(), text_1.split()))
    y = doc.add_paragraph()

    for i in range(0, len(diff)):
        f = len(diff) - 1
        if i < f:
            a = i - 1
        else:
            a = i

        if diff[i][0:3] == '  |':
            j = i + 1
            if diff[i][0:3] and diff[a][0:3] == '  |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '+ |':
            if diff[i][0:3] and diff[a][0:3] == '+ |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '- |':
            pass
        elif diff[i][0:3] == '  -':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0:3] == '  •':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0] == '+':
            w = len(diff[i])
            g = diff[i][1:w]
            y.add_run(g).font.color.rgb = RGBColor(255, 0, 0)
        elif diff[i][0] == '-':
            w = len(diff[i])
            g = diff[i][1:w]
            y.add_run(g).font.strike = True
        elif diff[i][0] == '?':
            pass
        else:
            if diff[i] != '+ |':
                y.add_run(diff[i])

    return doc

def compare_text_newandold(text_1, text_2, doc):
    '''
    function that places text into doc highlighting new and old text
    text_1: latest text. string.
    text_2: last text. string
    doc: word doc
    '''

    comp = difflib.Differ()
    diff = list(comp.compare(text_2.split(), text_1.split()))
    new_text = diff
    y = doc.add_paragraph()

    for i in range(0, len(diff)):
        f = len(diff) - 1
        if i < f:
            a = i - 1
        else:
            a = i

        if diff[i][0:3] == '  |':
            j = i + 1
            if diff[i][0:3] and diff[a][0:3] == '  |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '+ |':
            if diff[i][0:3] and diff[a][0:3] == '+ |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '- |':
            pass
        elif diff[i][0:3] == '  -':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0:3] == '  •':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0] == '+':
            w = len(diff[i])
            g = diff[i][1:w]
            y.add_run(g).font.color.rgb = RGBColor(255, 0, 0)
        elif diff[i][0] == '-':
            pass
        elif diff[i][0] == '?':
            pass
        else:
            if diff[i] != '+ |':
                y.add_run(diff[i][1:])

    return doc

def printing(project_name, milestone_master):
    '''function that compiles the summary sheet'''

    master_list = list_of_masters_all[0:4]

    doc = Document()
    print(project_name)
    heading = str(project_name)
    name = str(project_name)
    # TODO: change heading font size
    # todo be able to change text size and font
    intro = doc.add_heading(str(heading), 0)
    intro.alignment = 1
    intro.bold = True

    y = doc.add_paragraph()
    a = list_of_masters_all[0].data[project_name]['Senior Responsible Owner (SRO)']
    if a == None:
        a = 'TBC'
    else:
        a = a

    b = list_of_masters_all[0].data[project_name]['SRO Phone No.']
    if b == None:
        b = 'TBC'
    else:
        b = b

    y.add_run('SRO name:  ' + str(a) + ',   Tele:  ' + str(b))

    y = doc.add_paragraph()
    a = list_of_masters_all[0].data[project_name]['Project Director (PD)']
    if a == None:
        a = 'TBC'
    else:
        a = a
        b = list_of_masters_all[0].data[project_name]['PD Phone No.']
        if b == None:
            b = 'TBC'
        else:
            b = b

    y.add_run('PD name:  ' + str(a) + ',   Tele:  ' + str(b))

    '''Start of table with DCA confidence ratings'''
    table1 = doc.add_table(rows=1, cols=5)
    table1.cell(0, 0).width = Cm(7)

    '''quarter information in top row of table is here'''
    for i, quarter in enumerate(quarter_list):
        table1.cell(0, i+1).text = quarter

    # '''setting row height - partially working'''
    # # todo understand row height better
    # row = table1.rows[0]
    # tr = row._tr
    # trPr = tr.get_or_add_trPr()
    # trHeight = OxmlElement('w:trHeight')
    # trHeight.set(qn('w:val'), str(200))
    # trHeight.set(qn('w:hRule'), 'atLeast')
    # trPr.append(trHeight)

    SRO_conf_table_list = ['SRO DCA', 'Finance DCA', 'Benefits DCA', 'Resourcing DCA', 'Schedule DCA']
    SRO_conf_key_list = ['Departmental DCA', 'SRO Finance confidence', 'SRO Benefits RAG', 'Overall Resource DCA - Now',
                         'SRO Schedule Confidence']

    '''All SRO RAG rating placed in table'''
    for i in range(0, len(master_list)+1):
        table = doc.add_table(rows=1, cols=5)
        table.cell(0, 0).width = Cm(7)
        table.cell(0, 0).text = SRO_conf_table_list[i]
        for x, master in enumerate(master_list):
            try:
                rating = convert_rag_text(master.data[project_name][SRO_conf_key_list[i]])
                table.cell(0, x + 1).text = rating
                cell_colouring(table.cell(0, x + 1), rating)
            except (KeyError, TypeError):
                table.cell(0, x + 1).text = 'N/A'

    '''DCA Narrative text'''
    doc.add_paragraph()
    y = doc.add_paragraph()
    heading = 'SRO Overall DCA Narrative'
    y.add_run(str(heading)).bold = True

    dca_a = list_of_masters_all[0].data[project_name]['Departmental DCA Narrative']
    try:
        dca_b = list_of_masters_all[1].data[project_name]['Departmental DCA Narrative']
    except KeyError:
        dca_b = dca_a

    '''comparing text options'''
    # compare_text_showall(dca_a, dca_b, doc)
    compare_text_newandold(dca_a, dca_b, doc)

    '''Finance section'''
    y = doc.add_paragraph()
    heading = 'Financial information'
    y.add_run(str(heading)).bold = True

    '''Financial Meta data'''
    table1 = doc.add_table(rows=2, cols=5)
    table1.cell(0, 0).text = 'Forecast Whole Life Cost (£m):'
    table1.cell(0, 1).text = 'Percentage Spent:'
    table1.cell(0, 2).text = 'Source of Funding:'
    table1.cell(0, 3).text = 'Nominal or Real figures:'
    table1.cell(0, 4).text = 'Full profile reported:'

    wlc = round(list_of_masters_all[0].data[project_name]['Total Forecast'], 1)
    table1.cell(1, 0).text = str(wlc)
    # str(list_of_masters_all[0].data[project_name]['Total Forecast'])
    #a = list_of_masters_all[0].data[project_name]['Total Forecast']
    b = list_of_masters_all[0].data[project_name]['Pre 19-20 RDEL Forecast Total']
    if b == None:
        b = 0
    c = list_of_masters_all[0].data[project_name]['Pre 19-20 CDEL Forecast Total']
    if c == None:
        c = 0
    d = list_of_masters_all[0].data[project_name]['Pre 19-20 Forecast Non-Gov']
    if d == None:
        d = 0
    e = b + c + d
    try:
        c = round(e / wlc * 100, 1)
    except (ZeroDivisionError, TypeError):
        c = 0
    table1.cell(1, 1).text = str(c) + '%'
    a = str(list_of_masters_all[0].data[project_name]['Source of Finance'])
    b = list_of_masters_all[0].data[project_name]['Other Finance type Description']
    if b == None:
        table1.cell(1, 2).text = a
    else:
        table1.cell(1, 2).text = a + ' ' + str(b)
    table1.cell(1, 3).text = str(list_of_masters_all[0].data[project_name]['Real or Nominal - Actual/Forecast'])
    table1.cell(1, 4).text = ''

    '''Finance DCA Narrative text'''
    doc.add_paragraph()
    y = doc.add_paragraph()
    heading = 'SRO Finance Narrative'
    y.add_run(str(heading)).bold = True

    #TODO further testing on code down to 308. current hard code solution not ideal, plus not sure working properly yet

    gmpp_narrative_keys = ['Project Costs Narrative', 'Cost comparison with last quarters cost narrative',
                           'Cost comparison within this quarters cost narrative']

    fin_text_1 = combine_narrtives(project_name, list_of_masters_all[0], gmpp_narrative_keys)
    try:
        fin_text_2 = combine_narrtives(project_name, list_of_masters_all[1], gmpp_narrative_keys)
    except KeyError:
        fin_text_2 = fin_text_1

    # if narrative == 'NoneNoneNone':
    #     fin_text = combine_narrtives(name, dictionary_1, bicc_narrative_keys)
    # else:
    #     fin_text = narrative

    compare_text_newandold(fin_text_1, fin_text_2, doc)
    #compare_text_showall()

    '''financial chart heading'''
    y = doc.add_paragraph()
    heading = 'Financial Analysis - Cost Profile'
    y.add_run(str(heading)).bold = True
    y = doc.add_paragraph()
    y.add_run('{insert chart}')

    '''milestone section'''
    y = doc.add_paragraph()
    heading = 'Planning information'
    y.add_run(str(heading)).bold = True

    '''Milestone Meta data'''
    table1 = doc.add_table(rows=2, cols=4)
    table1.cell(0, 0).text = 'Project Start Date:'
    table1.cell(0, 1).text = 'Latest Approved Business Case:'
    table1.cell(0, 2).text = 'Start of Operations:'
    table1.cell(0, 3).text = 'Project End Date:'

    key_dates = milestone_master[project_name]

    #c = key_dates['Start of Project']
    try:
        c = tuple(key_dates['Start of Project'])[0]
        c = datetime.datetime.strptime(c.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
    except (KeyError, AttributeError):
        c = 'Not reported'

    table1.cell(1, 0).text = str(c)

    table1.cell(1, 1).text = str(list_of_masters_all[0].data[project_name]['BICC approval point'])

    try:
        a = tuple(key_dates['Start of Operation'])[0]
        a = datetime.datetime.strptime(a.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
        table1.cell(1, 2).text = str(a)
    except (KeyError, AttributeError):
        table1.cell(1, 2).text = 'Not reported'

    #b = key_dates['Project End Date']
    try:
        b = tuple(key_dates['Project End Date'])[0]
        b = datetime.datetime.strptime(b.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
    except (KeyError, AttributeError):
        b = 'Not reported'
    table1.cell(1, 3).text = str(b)

    # TODO: workout generally styling options for doc, paragraphs and tables

    '''milestone narrative text'''
    doc.add_paragraph()
    y = doc.add_paragraph()
    heading = 'SRO Milestone Narrative'
    y.add_run(str(heading)).bold = True

    mile_dca_a = list_of_masters_all[0].data[project_name]['Milestone Commentary']
    if mile_dca_a == None:
        mile_dca_a = 'None'

    try:
        mile_dca_b = list_of_masters_all[1].data[project_name]['Milestone Commentary']
        if mile_dca_b == None:
            mile_dca_b = 'None'
    except KeyError:
        mile_dca_b = mile_dca_a

    # compare_text_showall()
    compare_text_newandold(mile_dca_a, mile_dca_b, doc)

    '''milestone chart heading'''
    y = doc.add_paragraph()
    heading = 'Project reported high-level milestones and schedule changes'
    y.add_run(str(heading)).bold = True
    y = doc.add_paragraph()
    some_text = 'The below table presents all project reported remaining high-level milestones, with six months grace ' \
                'from close of the current quarter. Milestones are sorted in chronological order. Changes in milestones' \
                ' dates in comparison to last quarter and baseline have been calculated and are provided.'
    y.add_run(str(some_text)).italic = True
    y = doc.add_paragraph()
    y.add_run('{insert chart}')

    return doc

def combine_narrtives(project_name, master, key_list):
    '''function that combines text across different keys'''
    output = ''
    for key in key_list:
        output = output + str(master.data[project_name][key])

    return output


'''RUNNING PROGRAMME'''

quarter_list = ['This Qrt', 'Q2 1920', 'Q1 1920', 'Q4 1819']

'''One. select list of projects that dashboards should be built for'''
'''option one all'''
project_name_list = latest_quarter_project_names
'''option two select group - in dev'''
#projects = []
'''option three one project'''
#projects = ['High Speed Rail Programme (HS2)']

'''TWO. enter the project list variable into the below function. NOTE no change required. This is to ensure that the 
correct milestone dates are displayed in milestone meta data section'''
milestones = all_milestone_data_bulk(project_name_list, list_of_masters_all[0])

'''4) enter file path to where files should be saved. NOTE {} to be kept in file path as this is where project name is 
eventually placed in project file title'''
for project_name in project_name_list:
    a = printing(project_name, milestones)
    a.save(root_path/'output/proj_summaries/q3_1920_{}_summary.docx'.format(project_name))

