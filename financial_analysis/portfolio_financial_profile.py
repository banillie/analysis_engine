'''
Programme that creates a financial profiles for a group of projects i.e. can produce the portfolio profile or a chosen
set of projects profile.

Output document is an excel wb with three tabs containing the groups'latest quarter', 'last quarter', and
'baseline quarter' financial profiles

Just data handling for now. No graphics produced. To be done manually.

See instructions below.

Note: all master data is taken from the data file. Make sure this is up to date and that all relevant data is in
the import statement.
'''

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font
from analysis.data import q2_1920, cost_list, income_list, year_list, baseline_bcs, latest_cost_profiles, \
    last_cost_profiles, baseline_cost_profiles, latest_income_profiles, last_income_profiles, \
    baseline_income_profiles, dont_double_count
from analysis.engine_functions import filter_project_group, calculate_group_project_total

def place_in_excel(project_name_list):
    wb = Workbook()

    financial_profile_list = ['Latest Profile', 'Last quarter profile', 'Baseline Profile']
    cost_profile_data_list = [latest_cost_profiles, last_cost_profiles, baseline_cost_profiles]
    income_profile_data_list = [latest_income_profiles, last_income_profiles, baseline_income_profiles]

    for p, profile in enumerate(financial_profile_list):
        print(profile)
        '''worksheet is created for each project'''
        ws = wb.create_sheet(profile, p)  # creating worksheets
        ws.title = profile  # title of worksheet

        '''place information in each sheet'''
        ws.cell(row=1, column=1).value = 'Project'
        for i, project_name in enumerate(project_name_list):
            print(project_name)
            '''lists project names in row one'''
            ws.cell(row=1, column=i + 2).value = project_name

            '''iterates through financial dictionary - placing financial data in ws'''
            row_number = 1
            for cost in cost_list:
                for year in year_list:
                    ws.cell(row=row_number + 1, column=i+2).value = cost_profile_data_list[p][project_name][year + cost]
                    row_number += 1

        '''places totals in final column'''

        cost_totals = calculate_group_project_total(project_name_list, cost_profile_data_list[p], dont_double_count,
                                                    cost_list, year_list)

        ws.cell(row=1, column=len(project_name_list) + 2).value = 'Total'
        for i, value in enumerate(cost_totals.values()):
            ws.cell(row=i + 2, column=len(project_name_list)+2).value = value

        '''places keys into the chart in the first column'''
        for i, key in enumerate(cost_totals.keys()):
            ws.cell(row=i+2, column=1).value = key

        '''information on which projects are not included in totals'''
        ws.cell(row=1, column=len(project_name_list) + 4).value = 'Projects that have been removed to avoid double counting'
        for i, project in enumerate(dont_double_count):
            ws.cell(row=i + 2, column=len(project_name_list) + 4).value = project

        '''total cost data for output chart'''
        for z, year in enumerate(year_list):
            for x, type in enumerate(cost_list):
                ws.cell(row=z + 41, column=x + 2, value=cost_totals[year + type])

        '''labeling for total figure table '''
        ws.cell(row=40, column=1, value='Year')
        labeling_list_type = ['RDEL', 'CDEL', 'Non-Gov', 'Total']
        for i, label in enumerate(labeling_list_type):
            ws.cell(row=40, column=2 + i, value=label)

        '''labeling of years down the side'''
        for i, label in enumerate(year_list):
            ws.cell(row=41 + i, column=1, value=label)

        '''income total data'''
        income_totals = calculate_group_project_total(project_name_list, income_profile_data_list[p], dont_double_count,
                                                    income_list, year_list)

        for y, year in enumerate(year_list):
            for t, type in enumerate(income_list):
                ws.cell(row=y + 55, column=t + 2, value=income_totals[year + type])

        '''labeling for total figure table '''
        ws.cell(row=54, column=1, value='Year')
        ws.cell(row=54, column=2, value='Income Totals')

        '''labeling of years down the side'''
        for i, label in enumerate(year_list):
            ws.cell(row=55 + i, column=1, value=label)

    return wb


''' RUNNING PROGRAMME'''

'''ONE. set project name list options - this is where the group of projects is specified '''
'''option 1 - all '''
latest_quarter_projects = q2_1920.projects
'''option two - group of projects. use filter_project_group function'''
project_group_list = filter_project_group(q2_1920, 'HSMRPG')
'''option three - single project'''
one_project_list = []

'''TWO. place the variable containing the group of interest into the place_in_excel function and specify file path
to where output wb should be save'''
output = place_in_excel(project_group_list)
output.save("C:\\Users\\Standalone\\general\\masters folder\\portfolio_financial_profile\\"
            "q2_1920_portfolio_financial_profiles.xlsx")