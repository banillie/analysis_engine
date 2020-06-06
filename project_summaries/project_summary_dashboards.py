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
from docx.shared import Cm, RGBColor, Inches, Pt
from docx.enum.section import WD_SECTION_START, WD_ORIENT
import difflib
import matplotlib.pyplot as plt
import numpy as np

from analysis.engine_functions import convert_rag_text, all_milestone_data_bulk, project_time_difference, \
    milestone_analysis_date, all_milestones_dict
from analysis.data import list_of_masters_all, root_path, latest_cost_profiles, last_cost_profiles, \
    baseline_1_cost_profiles, year_list, SRO_conf_key_list, SRO_conf_table_list, bc_index, p_current_milestones, \
    p_last_milestones, first_diff_data, \
    a66, a303, crossrail, thameslink

import os

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

def produce_word_doc():
    '''Function that compiles each summary sheet'''

    master_list = list_of_masters_all[0:4]
    test_project_list = [a66]
                         #,a303, crossrail, thameslink]

    for project_name in test_project_list:
        doc = Document()
        print(project_name)

        font = doc.styles['Normal'].font
        font.name = 'Arial'
        font.size = Pt(10)

        heading = str(project_name)
        intro = doc.add_heading(str(heading), 0)
        intro.alignment = 1
        intro.bold = True

        para_1 = doc.add_paragraph()
        sro_name = list_of_masters_all[0].data[project_name]['Senior Responsible Owner (SRO)']
        if sro_name is None:
            sro_name = 'TBC'

        sro_phone = list_of_masters_all[0].data[project_name]['SRO Phone No.']
        if sro_phone == None:
            sro_phone = 'TBC'

        para_1.add_run('SRO name:  ' + str(sro_name) + ',   Tel:  ' + str(sro_phone))

        para_2 = doc.add_paragraph()
        pd_name = list_of_masters_all[0].data[project_name]['Project Director (PD)']
        if pd_name is None:
            pd_name = 'TBC'

        pd_phone = list_of_masters_all[0].data[project_name]['PD Phone No.']
        if pd_phone is None:
            pd_phone = 'TBC'

        para_2.add_run('PD name:  ' + str(pd_name) + ',   Tel:  ' + str(pd_phone))

        '''Start of table with DCA confidence ratings'''
        table = doc.add_table(rows=1, cols=5)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Delivery confidence'
        hdr_cells[1].text = 'This quarter'
        hdr_cells[2].text = str(list_of_masters_all[1].quarter)
        hdr_cells[3].text = str(list_of_masters_all[2].quarter)
        hdr_cells[4].text = str(list_of_masters_all[3].quarter)

        #RAG ratings
        for x, dca_key in enumerate(SRO_conf_key_list):
            row_cells = table.add_row().cells
            row_cells[0].text = dca_key
            for i, master in enumerate(master_list):
                try:
                    rating = convert_rag_text(master.data[project_name][dca_key])
                    row_cells[i+1].text = rating
                    cell_colouring(row_cells[i+1], rating)
                except (KeyError, TypeError):
                    row_cells[i].text = 'N/A'

        table.style = 'Table Grid'

        # TODO develop way of setting table column widths
        # TODO add text re what red text means.
        #set_col_widths(table)

        #DCA narrative
        doc.add_paragraph()

        headings_list = ['SRO delivery confidence narrative',
                         'Financial cost narrative',
                         'Financial comparison with last quarter',
                         'Financial comparison this quarter',
                         'Benefits Narrative',
                         'Benefits comparison with last quarter',
                         'Benefits comparison this quarter',
                         'Milestone narrative']

        narrative_keys_list = ['Departmental DCA Narrative',
                               'Project Costs Narrative',
                               'Cost comparison with last quarters cost narrative',
                               'Cost comparison within this quarters cost narrative',
                               'Benefits Narrative',
                               'Ben comparison with last quarters cost - narrative',
                               'Ben comparison within this quarters cost - narrative',
                               'Milestone Commentary']

        for x in range(len(headings_list)):
            doc.add_paragraph().add_run(str(headings_list[x])).bold = True
            text_one = str(list_of_masters_all[0].data[project_name][narrative_keys_list[x]])
            try:
                text_two = str(list_of_masters_all[1].data[project_name][narrative_keys_list[x]])
            except KeyError:
                text_two = text_one

            #different options for comparing costs
            # compare_text_showall(dca_a, dca_b, doc)
            compare_text_newandold(text_one, text_two, doc)

        #start of analysis
        new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
        new_width, new_height = new_section.page_height, new_section.page_width
        new_section.orientation = WD_ORIENT.LANDSCAPE
        new_section.page_width = new_width
        new_section.page_height = new_height

        #Finance chart
        #doc.add_paragraph().add_run(str('Financial Analysis')).bold = True #don't need this
        two_graph_chart(doc, project_name)

        #milestone table
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        #table heading
        doc.add_paragraph().add_run(str('Project high-level milestones')).bold = True
        #doc.add_paragraph()
        some_text = 'The below table presents all project reported remaining high-level milestones, with six months grace ' \
                    'from close of the current quarter. Milestones are sorted in chronological order. Changes in milestones' \
                    ' dates in comparison to last quarter and baseline have been calculated and are provided.'
        doc.add_paragraph().add_run(str(some_text)).italic = True

        milestone_table(doc, all_milestones_dict, project_name)


        # '''Financial Meta data'''
        # table1 = doc.add_table(rows=2, cols=5)
        # table1.cell(0, 0).text = 'Forecast Whole Life Cost (£m):'
        # table1.cell(0, 1).text = 'Percentage Spent:'
        # table1.cell(0, 2).text = 'Source of Funding:'
        # table1.cell(0, 3).text = 'Nominal or Real figures:'
        # table1.cell(0, 4).text = 'Full profile reported:'
        #
        # wlc = round(list_of_masters_all[0].data[project_name]['Total Forecast'], 1)
        # table1.cell(1, 0).text = str(wlc)
        # # str(list_of_masters_all[0].data[project_name]['Total Forecast'])
        # #a = list_of_masters_all[0].data[project_name]['Total Forecast']
        # b = list_of_masters_all[0].data[project_name]['Pre 19-20 RDEL Forecast Total']
        # if b == None:
        #     b = 0
        # c = list_of_masters_all[0].data[project_name]['Pre 19-20 CDEL Forecast Total']
        # if c == None:
        #     c = 0
        # d = list_of_masters_all[0].data[project_name]['Pre 19-20 Forecast Non-Gov']
        # if d == None:
        #     d = 0
        # e = b + c + d
        # try:
        #     c = round(e / wlc * 100, 1)
        # except (ZeroDivisionError, TypeError):
        #     c = 0
        # table1.cell(1, 1).text = str(c) + '%'
        # a = str(list_of_masters_all[0].data[project_name]['Source of Finance'])
        # b = list_of_masters_all[0].data[project_name]['Other Finance type Description']
        # if b == None:
        #     table1.cell(1, 2).text = a
        # else:
        #     table1.cell(1, 2).text = a + ' ' + str(b)
        # table1.cell(1, 3).text = str(list_of_masters_all[0].data[project_name]['Real or Nominal - Actual/Forecast'])
        # table1.cell(1, 4).text = ''



        #
        # '''milestone section'''
        # y = doc.add_paragraph()
        # heading = 'Planning information'
        # y.add_run(str(heading)).bold = True
        #
        # '''Milestone Meta data'''
        # table1 = doc.add_table(rows=2, cols=4)
        # table1.cell(0, 0).text = 'Project Start Date:'
        # table1.cell(0, 1).text = 'Latest Approved Business Case:'
        # table1.cell(0, 2).text = 'Start of Operations:'
        # table1.cell(0, 3).text = 'Project End Date:'
        #
        # key_dates = milestone_master[project_name]
        #
        # #c = key_dates['Start of Project']
        # try:
        #     c = tuple(key_dates['Start of Project'])[0]
        #     c = datetime.datetime.strptime(c.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
        # except (KeyError, AttributeError):
        #     c = 'Not reported'
        #
        # table1.cell(1, 0).text = str(c)
        #
        # table1.cell(1, 1).text = str(list_of_masters_all[0].data[project_name]['IPDC approval point'])
        #
        # try:
        #     a = tuple(key_dates['Start of Operation'])[0]
        #     a = datetime.datetime.strptime(a.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
        #     table1.cell(1, 2).text = str(a)
        # except (KeyError, AttributeError):
        #     table1.cell(1, 2).text = 'Not reported'
        #
        # #b = key_dates['Project End Date']
        # try:
        #     b = tuple(key_dates['Project End Date'])[0]
        #     b = datetime.datetime.strptime(b.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
        # except (KeyError, AttributeError):
        #     b = 'Not reported'
        # table1.cell(1, 3).text = str(b)
        #
        # # TODO: workout generally styling options for doc, paragraphs and tables
        #
        # '''milestone narrative text'''
        # doc.add_paragraph()
        # y = doc.add_paragraph()
        # heading = 'SRO Milestone Narrative'
        # y.add_run(str(heading)).bold = True
        #
        # mile_dca_a = list_of_masters_all[0].data[project_name]['Milestone Commentary']
        # if mile_dca_a == None:
        #     mile_dca_a = 'None'
        #
        # try:
        #     mile_dca_b = list_of_masters_all[1].data[project_name]['Milestone Commentary']
        #     if mile_dca_b == None:
        #         mile_dca_b = 'None'
        # except KeyError:
        #     mile_dca_b = mile_dca_a
        #
        # # compare_text_showall()
        # compare_text_newandold(mile_dca_a, mile_dca_b, doc)
        #
        # '''milestone chart heading'''
        # y = doc.add_paragraph()
        # heading = 'Project reported high-level milestones and schedule changes'
        # y.add_run(str(heading)).bold = True
        # y = doc.add_paragraph()
        # some_text = 'The below table presents all project reported remaining high-level milestones, with six months grace ' \
        #             'from close of the current quarter. Milestones are sorted in chronological order. Changes in milestones' \
        #             ' dates in comparison to last quarter and baseline have been calculated and are provided.'
        # y.add_run(str(some_text)).italic = True
        # y = doc.add_paragraph()
        # y.add_run('{insert chart}')

        #TODO add quarter info in title
        doc.save(root_path/'output/{}_summary.docx'.format(project_name))

def combine_narrtives(project_name, master, key_list):
    '''function that combines text across different keys'''
    output = ''
    for key in key_list:
        output = output + str(master.data[project_name][key])

    return output

amended_year_list = year_list[:-1]

def set_col_widths(table):
    widths = (Inches(2), Inches(1), Inches(1), Inches(1), Inches(1))
    for col in table.columns:
        print(col)
        for idx, width in enumerate(widths):
            col.cells[idx].width = width

def get_financial_data(project_name, cost_type):
    '''gets project financial data'''
    latest = []
    last = []
    baseline = []
    for year in amended_year_list:
        baseline.append(baseline_1_cost_profiles[project_name][year + cost_type])
        last.append(last_cost_profiles[project_name][year + cost_type])
        latest.append(latest_cost_profiles[project_name][year + cost_type])

    return latest, last, baseline

def two_graph_chart(doc, project_name):

    #cost profile time change chart
    profile_data_total = get_financial_data(project_name, ' total')
    profile_data_rdel = get_financial_data(project_name, ' RDEL Forecast Total')
    profile_data_cdel = get_financial_data(project_name, ' CDEL Forecast Total')

    year = ['19/20', '20/21', '21/22', '22/23', '23/24', '24/25', '25/26', '26/27', '27/28', '28/29']
    baseline_profile_total = profile_data_total[2]
    last_profile_total = profile_data_total[1]
    latest_profile_total = profile_data_total[0]

    latest_profile_rdel = profile_data_rdel[0]
    latest_profile_cdel = profile_data_cdel[0]

    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2)

    fig.suptitle('Financial and benefits analysis', fontweight='bold') # title
    # set fig size
    fig.set_figheight(6)
    fig.set_figwidth(8)
    #fig.suptitle(str(project_name) + ' Cost Analysis')  # title

    ax1.plot(year, baseline_profile_total, color='#90c978', label='Baseline', linewidth=3.0, marker="o")
    ax1.plot(year, last_profile_total, color='#afd5aa', label='Last quarter', linewidth=3.0, marker="o")
    ax1.plot(year, latest_profile_total, color='#83c6dd', label='Latest', linewidth=3.0, marker="o")

    ax3.plot(year, latest_profile_cdel, color='#e2a9be', label='CDEL', linewidth=3.0, marker="o")
    ax3.plot(year, latest_profile_rdel, color='#c2d5a7', label='RDEL', linewidth=3.0, marker="o")

    # ax1.get_xaxis().set_visible(False) # hide x labels, but this hides the x grid lines also
    ax1.tick_params(axis='x', which='major', labelsize=6)
    # ax1.set_xlabel('Financial Years')
    ax1.set_ylabel('Cost (£m)')
    # xlab1 = ax1.xaxis.get_label()
    ylab1 = ax1.yaxis.get_label()
    # xlab1.set_style('italic')
    # xlab1.set_size(8)
    ylab1.set_style('italic')
    ylab1.set_size(8)
    ax1.grid(color='grey', linestyle='-', linewidth=0.2)
    ax1.legend(prop={'size': 8})
    ax1.set_title('Fig 1 - cost profile changes', loc='left', fontsize=8, fontweight='bold')

    ax3.tick_params(axis='x', which='major', labelsize=6)
    ax3.set_xlabel('Financial Years')
    ax3.set_ylabel('Cost (£m)')
    xlab3 = ax3.xaxis.get_label()
    ylab3 = ax3.yaxis.get_label()
    xlab3.set_style('italic')
    xlab3.set_size(8)
    ylab3.set_style('italic')
    ylab3.set_size(8)
    ax3.grid(color='grey', linestyle='-', linewidth=0.2)
    ax3.legend(prop={'size': 8})
    ax3.set_title('Fig 2 - cost profile spend type', loc='left', fontsize=8, fontweight='bold')


    #Total cost bar chart
    bar_chart_data = bar_chart_financial(project_name)
    profile_1 = bar_chart_data[1] - (bar_chart_data[0] + bar_chart_data[2])
    profile_2 = bar_chart_data[4] - (bar_chart_data[3] + bar_chart_data[5])
    profile_3 = bar_chart_data[7] - (bar_chart_data[6] + bar_chart_data[8])

    labels = ['Baseline', 'Last Quarter', 'Latest']
    spent = np.array([bar_chart_data[0],
                      bar_chart_data[3],
                      bar_chart_data[6]])
    profiled = np.array([profile_1,
                        profile_2,
                         profile_3])
    unprofiled = np.array([bar_chart_data[2],
                           bar_chart_data[5],
                           bar_chart_data[8]])
    width = 0.45
    ax2.bar(labels, spent, width, label='Spent', color='#F7DB4F')
    ax2.bar(labels, profiled, width, bottom=spent, label='Profiled', color='#F26B38')
    ax2.bar(labels, unprofiled, width, bottom=spent+profiled, label='Unprofiled', color='#EC2049')
    ax2.legend(prop={'size': 6})
    ax2.set_ylabel('Cost (£m)')
    ylab2 = ax2.yaxis.get_label()
    ylab2.set_style('italic')
    ylab2.set_size(8)
    ax2.tick_params(axis='x', which='major', labelsize=6)
    ax2.tick_params(axis='y', which='major', labelsize=6)
    ax2.set_title('Fig 3 - wlc break down', loc='left', fontsize=8, fontweight='bold')

    #Total benefits Bar Chart
    bar_chart_data = bar_chart_benefits(project_name)
    profile_1 = bar_chart_data[1] - (bar_chart_data[0] + bar_chart_data[2])
    profile_2 = bar_chart_data[4] - (bar_chart_data[3] + bar_chart_data[5])
    profile_3 = bar_chart_data[7] - (bar_chart_data[6] + bar_chart_data[8])

    labels = ['Baseline', 'Last Quarter', 'Latest']
    spent = np.array([bar_chart_data[0],
                      bar_chart_data[3],
                      bar_chart_data[6]])
    profiled = np.array([profile_1,
                         profile_2,
                         profile_3])
    unprofiled = np.array([bar_chart_data[2],
                           bar_chart_data[5],
                           bar_chart_data[8]])
    width = 0.5
    ax4.bar(labels, spent, width, label='Delivered', color='#99B898')
    ax4.bar(labels, profiled, width, bottom=spent, label='Profiled', color='#FECEA8')
    ax4.bar(labels, unprofiled, width, bottom=spent + profiled, label='Unprofiled', color='#FF847C')
    ax4.legend(prop={'size': 6})
    ax4.set_ylabel('Benefits (£m)')
    ylab3 = ax4.yaxis.get_label()
    ylab3.set_style('italic')
    ylab3.set_size(8)
    ax4.tick_params(axis='x', which='major', labelsize=6)
    ax4.tick_params(axis='y', which='major', labelsize=6)
    ax4.set_title('Fig 4 - whole life benefits break down', loc='left', fontsize=8, fontweight='bold')

    fig.savefig('cost_profile.png')

    doc.add_picture('cost_profile.png', width=Inches(8))  # to place nicely in doc

    return doc

def milestone_table(doc, function, project_name):

    p_oldest_milestones = function([project_name], list_of_masters_all[bc_index[project_name][2]])
    second_diff_data = project_time_difference(p_current_milestones, p_oldest_milestones)

    table = doc.add_table(rows=1, cols=5)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Milestone'
    hdr_cells[1].text = 'Date'
    hdr_cells[2].text = 'Change from Lst Qrt'
    hdr_cells[3].text = 'Change from BL'
    hdr_cells[4].text = 'Notes'

    # TODO specify column widths

    for milestone in p_current_milestones[project_name].keys():

        milestone_date = tuple(p_current_milestones[project_name][milestone])[0]

        try:
            if milestone_analysis_date <= milestone_date: # filter based on date
                row_cells = table.add_row().cells
                row_cells[0].text = milestone
                if milestone_date is None:
                    row_cells[1].text = 'No date'
                else:
                    row_cells[1].text = milestone_date.strftime("%d/%m/%Y")
                b_one_value = first_diff_data[project_name][milestone]
                row_cells[2].text = plus_minus_days(b_one_value)
                b_two_value = second_diff_data[project_name][milestone]
                row_cells[3].text = plus_minus_days(b_two_value)
                notes = p_current_milestones[project_name][milestone][milestone_date]
                row_cells[4].text = str(notes)
                paragraph = row_cells[4].paragraphs[0]
                run = paragraph.runs
                font = run[0].font
                font.size = Pt(8)  # font size = 8


        except TypeError:  # this is to deal with none types which are still placed in output
            row_cells = table.add_row().cells
            row_cells[0].text = milestone
            if milestone_date is None:
                row_cells[1].text = 'No date'
            else:
                row_cells[1].text = milestone_date.strftime("%d/%m/%Y")
            b_one_value = first_diff_data[project_name][milestone]
            row_cells[2].text = plus_minus_days(b_one_value)
            b_two_value = second_diff_data[project_name][milestone]
            row_cells[3].text = plus_minus_days(b_two_value)
            notes = p_current_milestones[project_name][milestone][milestone_date]
            row_cells[4].text = str(notes)
            paragraph = row_cells[4].paragraphs[0]
            run = paragraph.runs
            font = run[0].font
            font.size = Pt(8)  # font size = 8

    table.style = 'Table Grid'

    make_rows_bold(table.rows[0]) # makes top of table bold. Found function on stack overflow.

    return doc

def make_rows_bold(*rows):
    '''Makes text bold in specified row'''
    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True

def plus_minus_days(change_value):
    '''mini function to place plus or minus sign before time delta
    value in milestone_table function'''
    try:
        if change_value > 0:
            text = '+ ' + str(change_value)
        # elif change_value < 0:
        #     text = '- ' + str(change_value)
        else:
            text = str(change_value)
    except TypeError:
        text = change_value

    return text

def bar_chart_financial(project_name):
    '''gets financial data to place into the bar chart element in the financial analysis graphs'''
    key_list = [('Pre-profile RDEL Forecast Total',
                'Pre-profile CDEL Forecast Total'),
                ('Total RDEL Forecast Total',
                 'Total CDEL Forecast Total WLC'),
                ('Unprofiled RDEL Forecast Total',
                 'Unprofiled CDEL Forecast Total WLC')]

    cost_list = []

    index_1 = bc_index[project_name]
    index_2 = index_1[0:3]
    index_2.reverse()
    for x in index_2:
        for y in key_list:
            rdel = list_of_masters_all[x].data[project_name][y[0]]
            cdel = list_of_masters_all[x].data[project_name][y[1]]
            total = rdel + cdel
            cost_list.append(total)

    return cost_list

def bar_chart_benefits(project_name):
    '''gets benefits data to place into the bar chart element in the financial analysis graphs'''

    key_list = ['Pre-profile BEN Total',
                'Unprofiled Remainder BEN Forecast - Total Monetised Benefits',
                'Total BEN Forecast - Total Monetised Benefits']

    cost_list = []

    index_1 = bc_index[project_name]
    index_2 = index_1[0:3]
    index_2.reverse()
    for x in index_2:
        print(x)
        for y in key_list:
            ben = list_of_masters_all[x].data[project_name][y]
            print(ben)
            cost_list.append(ben)

    return cost_list

'''RUNNING PROGRAMME'''

'''enter into the printing function the quarter details for the output files e.g. _q4_1920 (note put underscore at 
front)'''
produce_word_doc()

#remove the cost profile document created while programme is running
os.remove('/home/will/code/python/analysis_engine/project_summaries/cost_profile.png')