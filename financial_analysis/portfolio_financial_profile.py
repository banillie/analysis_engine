'''
Programme to create a financial profile for a group of projects i.e. can produce the portfolio profile or a chosen
set of projects profile.

Output document is an excel spreadsheet contain a graph with financial profile

See instructions below.

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.


'''

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font
from analysis.data import financial_analysis_masters_list, q2_1920
from analysis.engine_functions import bc_ref_stages, master_baseline_index, filter_project_group



def year_totals(project_list, remove_from_totals_list, data_key_list, q_masters_dict_list, q_masters_list, period):
    #TODO convert output into dictionary

    fy_total_list = []

    totals_project_list = []
    for project_name in project_list:
        if project_name not in remove_from_totals_list:
            totals_project_list.append(project_name)

    for key in data_key_list:
        thesum = 0
        for project_name in totals_project_list:
            try:
                if period == 'baseline':
                    to_add = q_masters_dict_list[q_masters_list[project_name][2]].data[project_name][key]
                if period == 'last':
                    to_add = q_masters_dict_list[q_masters_list[project_name][1]].data[project_name][key]
                if period == 'latest':
                    to_add = q_masters_dict_list[q_masters_list[project_name][0]].data[project_name][key]
                thesum = thesum + to_add
            except (TypeError, KeyError):
                pass
        fy_total_list.append(thesum)

    return fy_total_list

# def likeforlike():
#     '''
#     small programme used to filter out projects that are not in both data sets
#     :param data_1: most recent quarters data
#     :param data_2: less recent quarters data
#     :return: a list of projects that are in both data sets
#     '''
#
#     one = list(set(q1_1920) - set(q4_1819))
#     two = list(set(q4_1819) - set(q1_1920))
#
#     output_list = one + two
#
#     return output_list

def place_in_excel(project_list, data_key_list, total_data, q_masters_dict_list, q_masters_list, period):
    wb = Workbook()
    ws = wb.active

    ws.cell(row=1, column=1).value = 'Project'
    for i, project_name in enumerate(project_list):
        '''lists project names in row one'''
        ws.cell(row=1, column=i + 2).value = project_name

        '''iterates through financial dictionary - placing financial data in ws'''
        for x, key in enumerate(data_key_list):
            try:
                if period == 'baseline':
                    ws.cell(row=x+2, column=i+2).value = \
                        q_masters_dict_list[q_masters_list[project_name][2]].data[project_name][key]
                if period == 'last':
                    ws.cell(row=x + 2, column=i + 2).value = \
                    q_masters_dict_list[q_masters_list[project_name][1]].data[project_name][key]
                if period == 'latest':
                    ws.cell(row=x + 2, column=i + 2).value = \
                    q_masters_dict_list[q_masters_list[project_name][0]].data[project_name][key]
            except KeyError:
                ws.cell(row=x + 2, column=i + 2).value = 0

    '''places totals in final column. to note because this is a list and not a dictionary as for fin_data there is 
    possibility that data could become unaligned. Whether changing the list of cells_to_capture causes them to become
    unaligned needs to be tested'''
    ws.cell(row=1, column=len(project_list) + 2).value = 'Total'
    for i, values in enumerate(total_data):
        ws.cell(row=i + 2, column=len(project_list)+2).value = values

    '''places keys into the chart in the first column'''
    for i, key in enumerate(data_key_list):
        ws.cell(row=i+2, column=1).value = key

    '''information on which projects are not included in totals'''
    ws.cell(row=1, column=len(project_list) + 4).value = 'Projects that have been removed to avoid double counting'
    for i, project in enumerate(dont_double_count):
        ws.cell(row=i + 2, column=len(project_list) + 4).value = project

    '''data for overall chart. As above because this data is in a list - possibility of it being unaligned needs 
    testing. not the best way of managing data flow, but working for now'''
    start_row = len(total_data) + 8
    for x in range(0, int(len(total_data) / 4)):
        ws.cell(row=start_row, column=2, value=total_data[x])
        start_row += 1

    start_row = len(total_data) + 8
    for x in range(int(len(total_data) / 4), (int(len(total_data) / 4) * 2)):
        ws.cell(row=start_row, column=3, value=total_data[x])
        start_row += 1

    start_row = len(total_data) + 8
    for x in range((int(len(total_data) / 4) * 2), (int(len(total_data) / 4) * 3)):
        ws.cell(row=start_row, column=4, value=total_data[x])
        start_row += 1

    start_row = len(total_data) + 8
    for x in range((int(len(total_data) / 4) * 3), int(len(total_data))):
        ws.cell(row=start_row, column=5, value=total_data[x])
        start_row += 1

    '''code was essentially a hack'''

    start_row = len(total_data) + 8
    list_of_numbers = [0, len(capture_rdel), len(capture_rdel)*2]
    total_sum = 0
    for i in range(0, len(capture_rdel)):
        for x in list_of_numbers:
            total_sum = total_sum + total_data[x + i]
            ws.cell(row=start_row, column=6, value=total_sum)
        start_row += 1
        total_sum = 0

    a = len(total_data) + 7
    ws.cell(row=a, column=2, value='RDEL')
    ws.cell(row=a, column=3, value='CDEL')
    ws.cell(row=a, column=4, value='Non-Gov')
    ws.cell(row=a, column=5, value='Income')
    ws.cell(row=a, column=6, value='Total')


    # ws.cell(row=a+1, column=1, value='17/18')
    #ws.cell(row=a + 1, column=1, value='18/19')
    ws.cell(row=a + 1, column=1, value='19/20')
    ws.cell(row=a + 2, column=1, value='20/21')
    ws.cell(row=a + 3, column=1, value='21/22')
    ws.cell(row=a + 4, column=1, value='22/23')
    ws.cell(row=a + 5, column=1, value='23/24')
    ws.cell(row=a + 6, column=1, value='24/25')
    ws.cell(row=a + 7, column=1, value='25/26')
    ws.cell(row=a + 8, column=1, value='26/27')
    ws.cell(row=a + 9, column=1, value='27/28')
    ws.cell(row=a + 10, column=1, value='28/29')
    ws.cell(row=a + 11, column=1, value='Unprofiled')

    '''this builds a very basic chart'''
    # TODO fix chart
    chart = LineChart()
    chart.title = 'Portfolio cost profile'
    chart.style = 4
    chart.x_axis.title = 'Financial Year'
    chart.y_axis.title = 'Cost (£m)'
    chart.height = 15  # default is 7.5
    chart.width = 26  # default is 15

    '''styling chart'''
    # axis titles
    font = Font(typeface='Calibri')
    size = 1200  # 12 point size
    cp = CharacterProperties(latin=font, sz=size, b=True)  # Bold
    pp = ParagraphProperties(defRPr=cp)
    rtp = RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])
    chart.x_axis.title.tx.rich.p[0].pPr = pp
    chart.y_axis.title.tx.rich.p[0].pPr = pp

    # title
    size_2 = 1400
    cp_2 = CharacterProperties(latin=font, sz=size_2, b=True)
    pp_2 = ParagraphProperties(defRPr=cp_2)
    rtp_2 = RichText(p=[Paragraph(pPr=pp_2, endParaRPr=cp_2)])
    chart.title.tx.rich.p[0].pPr = pp_2

    data = Reference(ws, min_col=2, min_row=51, max_col=5, max_row=61)
    cats = Reference(ws, min_col=1, min_row=52, max_row=61)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    s3 = chart.series[0]
    s3.graphicalProperties.line.solidFill = "36708a"  # dark blue
    s8 = chart.series[1]
    s8.graphicalProperties.line.solidFill = "68db8b"  # green
    s9 = chart.series[2]
    s9.graphicalProperties.line.solidFill = "794747"  # dark red
    s9 = chart.series[3]
    s9.graphicalProperties.line.solidFill = "73527f"  # purple

    ws.add_chart(chart, "I52")

    return wb

def run_financials_all(project_list, project_list_remove, data_key_list, q_masters_dict_list, period):
    baseline_bc = bc_ref_stages(project_list, q_masters_dict_list)
    q_masters_list = master_baseline_index(project_list, q_masters_dict_list, baseline_bc)
    total_data = year_totals(project_list, project_list_remove, data_key_list, q_masters_dict_list, q_masters_list, period)
    output = place_in_excel(project_list, data_key_list, total_data, q_masters_dict_list, q_masters_list, period)

    return output

'''List of financial data keys to capture. This should be amended to years of interest'''
capture_rdel = ['19-20 RDEL Forecast Total', '20-21 RDEL Forecast Total', '21-22 RDEL Forecast Total',
                '22-23 RDEL Forecast Total', '23-24 RDEL Forecast Total', '24-25 RDEL Forecast Total',
                '25-26 RDEL Forecast Total', '26-27 RDEL Forecast Total', '27-28 RDEL Forecast Total',
                '28-29 RDEL Forecast Total', 'Unprofiled RDEL Forecast Total']
capture_cdel = ['19-20 CDEL Forecast Total', '20-21 CDEL Forecast Total', '21-22 CDEL Forecast Total',
                 '22-23 CDEL Forecast Total', '23-24 CDEL Forecast Total', '24-25 CDEL Forecast Total',
                 '25-26 CDEL Forecast Total', '26-27 CDEL Forecast Total', '27-28 CDEL Forecast Total',
                 '28-29 CDEL Forecast Total', 'Unprofiled CDEL Forecast Total']
capture_ng = ['19-20 Forecast Non-Gov', '20-21 Forecast Non-Gov', '21-22 Forecast Non-Gov', '22-23 Forecast Non-Gov',
              '23-24 Forecast Non-Gov', '24-25 Forecast Non-Gov', '25-26 Forecast Non-Gov', '26-27 Forecast Non-Gov',
              '27-28 Forecast Non-Gov', '28-29 Forecast Non-Gov', 'Unprofiled Forecast-Gov']
capture_income =['19-20 Forecast - Income both Revenue and Capital',
                '20-21 Forecast - Income both Revenue and Capital', '21-22 Forecast - Income both Revenue and Capital',
                '22-23 Forecast - Income both Revenue and Capital', '23-24 Forecast - Income both Revenue and Capital',
                '24-25 Forecast - Income both Revenue and Capital', '25-26 Forecast - Income both Revenue and Capital',
                '26-27 Forecast - Income both Revenue and Capital', '27-28 Forecast - Income both Revenue and Capital',
                '28-29 Forecast - Income both Revenue and Capital', 'Unprofiled Forecast Income']
all_data_lists = capture_rdel + capture_cdel + capture_ng + capture_income

''' RUNNING PROGRAMME'''

'''ONE. Project name list options - this is where the group of projects is specified '''
'''option 1 - all '''
latest_quarter_projects = q2_1920.projects
'''option two - group of projects. use filter_project_group function'''
project_group_list = filter_project_group(q2_1920, 'HSMRPG')
'''option three - single project'''
one_project_list = ['High Speed Rail Programme (HS2)']

'''remove projects to stop double counting - this can normally be ignored unless a new project starts reporting 
that needs to be removed from double counting or counting in general
NOTE. option here to include a project if you want to see what the profile looks like with it e.g. Hs2 Programme'''
dont_double_count = ['HS2 Phase 2b', 'HS2 Phase1', 'HS2 Phase2a', 'East Midlands Franchise',
                     'West Coast Partnership Franchise']

'''TWO. Enter variables created via options above into function and run programme. The function is structured as 
follows... run_financials_all(project_list, project_list_remove, data_key_list, q_masters_dict_list)

1) project_list = list of projects to include in analysis. 
2) project_list_remove = list of projects to not be included in total figures. 
3) data_key_list = the list of financial keys. This will normally not change
4) q_masters_dict_list = the list of master dictionaries to include in analysis. This will normally not change. 
5) period = which financial information you want to return. options are 'baseline', 'last', 'latest' '''

output = run_financials_all\
    (latest_quarter_projects, dont_double_count, all_data_lists, financial_analysis_masters_list, 'last')

'''5) specify the file path to where the output file should be saved'''
output.save("C:\\Users\\Standalone\\general\\masters folder\\portfolio_financial_profile\\"
            "q2_1920_portfolio_financial_profile_last.xlsx")