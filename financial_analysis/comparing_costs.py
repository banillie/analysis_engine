'''
Programme to compare a specific year or total wlc information. It produces a wb output with data and calculations
only i.e. no graph. The output shows changes to wlc in relation 1) overall figures, 2) change between quarters,
3) percent change are highlighted in red if change is greater/less than £100m/-£100m or percentage change greater/less
than 5%/-5% of project value

It is from the data placed into the output document that a simple bard chart can be built to show the most significant
changes in cost since the previous quarter.
'''

from openpyxl import Workbook
from openpyxl.styles import Font
from analysis.engine_functions import bc_ref_stages, master_baseline_index
from analysis.data import financial_analysis_masters_list, q2_1920

def compare(data_1, data_2):
    wb = Workbook()
    ws = wb.active

    for i, name in enumerate(data_1):
        '''place project names into ws'''
        ws.cell(row=i+2, column=1).value = name

        '''loop for placing wlc data into ws. highlight changes between quarters in red'''
        latest_wlc = data_1[name]
        try:
            last_wlc = data_2[name]
        except KeyError:
            last_wlc = 'None'

        ws.cell(row=i + 2, column=2).value = latest_wlc

        if latest_wlc != last_wlc:
            ws.cell(row=i + 2, column=2).font = red_text

        if name in data_2.keys():
            try:
                ws.cell(row=i + 2, column=3).value = last_wlc
                change = latest_wlc - last_wlc
                if last_wlc > 0:
                    percent_change = (latest_wlc - last_wlc)/last_wlc
                else:
                    percent_change = (latest_wlc - last_wlc)/(last_wlc + 1)
                ws.cell(row=i + 2, column=4).value = change
                ws.cell(row=i + 2, column=5).value = percent_change
                if change >= 100 or change <= -100:
                    ws.cell(row=i + 2, column=4).font = red_text
                if percent_change >= 0.05 or percent_change <= -0.05:
                    ws.cell(row=i + 2, column=5).font = red_text
            except TypeError:
                pass
        else:
            ws.cell(row=i + 2, column=3).value = last_wlc

    ws.cell(row=1, column=1).value = 'Project Name'
    ws.cell(row=1, column=2).value = 'Latest Quarter'
    ws.cell(row=1, column=3).value = 'Baseline Quarter'
    ws.cell(row=1, column=4).value = 'Change'
    ws.cell(row=1, column=5).value = 'Percentage Change'
    return wb

def get_yearly_costs(proj_list, q_masters_dict_list, cost_list, year, index):
    output_dict = {}
    for proj_name in proj_list:
        project_dict = q_masters_dict_list[q_masters_list[proj_name][index]].data[proj_name]
        total = 0
        for type in cost_list:
            if year + type in project_dict.keys():
                cost = project_dict[year + type]
                try:
                    total = total + cost
                except TypeError:
                    pass

        output_dict[proj_name] = total

    return output_dict

def get_wlc(proj_list, q_masters_dict_list, key, index):
    output_dict = {}
    for proj_name in proj_list:
        proj_dict = q_masters_dict_list[q_masters_list[proj_name][index]].data[proj_name]
        total = proj_dict[key]
        output_dict[proj_name] = total

    return output_dict

red_text = Font(color="FF0000")

'''INSTRUCTIONS FOR RUNNING PROGRAMME'''

'''ONE. list of projects'''
proj_names_all = q2_1920.projects

baseline_bc = bc_ref_stages(proj_names_all, financial_analysis_masters_list)
q_masters_list = master_baseline_index(proj_names_all, financial_analysis_masters_list, baseline_bc)

'''TWO. decide which output you require'''

'''OPTION ONE - in year costs. only real action here is to adjust year_interest'''

'''in year cost lists is chosen through the cost list. No not change.'''
cost_list = [' RDEL Forecast Total', ' CDEL Forecast Total', ' Forecast Non-Gov']

'''in year income list is chosen through the income list. do not change. Note. not really in use'''
income_list = [' Forecast - Income both Revenue and Capital']

'''chose financial year of interest. change accordingly. needs to be in format of YY-YY'''
year_interest = '23-24'

'''get fy information by entering the appropriate variables'''
latest_fy = get_yearly_costs(proj_names_all, financial_analysis_masters_list, cost_list, year_interest, 0)
baseline_fy = get_yearly_costs(proj_names_all, financial_analysis_masters_list, cost_list, year_interest, 2)

'''OPTION TWO - for wlc costs. no real action required'''

'''chose wlc cost key of interest from master data. Get information by entering appropriate variables below'''
wlc_key = 'Total Forecast'
latest_wlc = get_wlc(proj_names_all, financial_analysis_masters_list, wlc_key, 0)
baseline_wlc = get_wlc(proj_names_all, financial_analysis_masters_list, wlc_key, 2)

'''THREE. enter desired variables into the compare function i.e. enter either one_fy, two_fy or one_wlc, two_wlc and 
specify file path for where output document to be saved'''
output = compare(latest_wlc, baseline_wlc)

output.save("C:\\Users\\Standalone\\general\\masters folder\\portfolio_financial_profile\\"
            "q2_1920_comparing_costs_wlc.xlsx")