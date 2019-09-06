'''

This programme creates a master spreadsheet for commissioning templates.

Documents required to run the programme are set out below. The latest versions of these should be taken from TiME
and saved onto laptops in the file paths at the bottom of the programme.

Documents required are:
1) the internal commission datamap (make sure you have the latest/final version).
2) latest quarter DfT_master spreadsheet. (Note in the commission the latest quarter data is used to both populate the
commission template in active reporting fields, as well as provide a record of what is reported last quarter.)

2) and 3) above are taken from the data module

'''

from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from analysis.data import q1_1920


def create_master(workbook, latest_q_data):
    ws = workbook.active
    type_list_2 = ['RDEL', 'CDEL', 'Non-Gov', 'Income', 'BEN'] # list of cost/ben types. used to remove none value entries

    red_text = Font(color="00fc2525")

    for i, name in enumerate(list(latest_q_data.keys())):
        print(name)
        ws.cell(row=1, column=2+i).value = name  # place project names in file

        # for loop for placing data into the worksheet
        for row_num in range(2, ws.max_row+1):
            key = ws.cell(row=row_num, column=1).value

            if key in latest_q_data[name].keys():
                ws.cell(row=row_num, column=2+i).value = latest_q_data[name][key]
            if "lst qrt" in key:
                altered_lst_quarter = key.replace("lst qrt ", "")
                try:
                    ws.cell(row=row_num, column=2 + i).value = latest_q_data[name][altered_lst_quarter]

                    '''handling of DN types'''
                    if key == "lst qrt DN Type 1":
                        ws.cell(row=row_num, column=2 + i).value = dn_combine_text(latest_q_data[name],['DN Type 1', 'DN Description 1'])
                    if key == "lst qrt DN Type 2":
                        ws.cell(row=row_num, column=2 + i).value = dn_combine_text(latest_q_data[name],['DN Type 2', 'DN Description 2'])
                    if key == "lst qrt DN Type 3":
                        ws.cell(row=row_num, column=2 + i).value = dn_combine_text(latest_q_data[name],['DN Type 3', 'DN Description 3'])
                    if key == "lst qrt DN Type 4":
                        ws.cell(row=row_num, column=2 + i).value = dn_combine_text(latest_q_data[name],['DN Type 4', 'DN Description 4'])
                    if key == "lst qrt DN Type 5":
                        ws.cell(row=row_num, column=2 + i).value = dn_combine_text(latest_q_data[name],['DN Type 5', 'DN Description 5'])

                    '''handling of strategic outcomes'''
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 1)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 1)',
                                                                        'IO1', 'IO1 - Monetised?', 'IO1 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 2)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 2)',
                                                                        'IO2', 'IO2 - Monetised?', 'IO2 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 3)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 3)',
                                                                        'IO3', 'IO3 - Monetised?', 'IO3 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 4)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 4)',
                                                                        'IO4', 'IO4 - Monetised?', 'IO4 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 5)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 5)',
                                                                        'IO5', 'IO5 - Monetised?', 'IO5 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 6)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 6)',
                                                                        'IO6', 'IO6 - Monetised?', 'IO6 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 7)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 7)',
                                                                        'IO7', 'IO7 - Monetised?', 'IO7 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 8)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 8)',
                                                                        'IO8', 'IO8 - Monetised?', 'IO8 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 8)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 8)',
                                                                        'IO8', 'IO8 - Monetised?', 'IO8 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 9)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 9)',
                                                                        'IO9', 'IO9 - Monetised?', 'IO9 PESTLE'])
                    if key == "lst qrt List Strategic Outcomes (GMPP - Intended Outcome 10)":
                        ws.cell(row=row_num, column=2 + i).value = \
                            strategic_combine_text(latest_q_data[name], ['List Strategic Outcomes (GMPP - Intended Outcome 10)',
                                                                        'IO10', 'IO10 - Monetised?', 'IO10 PESTLE'])

                    '''handling of investment objectives'''
                    if key == "lst qrt Primary investment Objective":
                        ws.cell(row=row_num, column=2 + i).value = dn_combine_text(latest_q_data[name],
                                                   ['Primary investment Objective', 'IO11 Monetised'])
                    if key == "lst qrt Secondary investment Objective":
                        ws.cell(row=row_num, column=2 + i).value = dn_combine_text(latest_q_data[name],
                                                   ['Secondary investment Objective', 'IO12 Monetised'])

                    '''handling of risk descriptions'''
                    if key == "lst qrt Brief Risk Decription 1":
                        ws.cell(row=row_num, column=2 + i).value = \
                            risk_combine_text(latest_q_data[name], [
                                'Brief Risk Decription 1', 'BRD 1Risk Category', 'BRD 1 Primary Risk to',
                                'BRD 1 Internal Control', 'BRD 1 Residual Impact', 'BRD 1 Residual Likelihood'])
                    if key == "lst qrt Brief Risk Decription 2":
                        ws.cell(row=row_num, column=2 + i).value = \
                            risk_combine_text(latest_q_data[name], [
                                'Brief Risk Decription 2', 'BRD 2 Risk Category', 'BRD 2 Primary Risk to',
                                'BRD 2 Internal Control', 'BRD 2 Residual Impact', 'BRD 2 Residual Likelihood'])
                    if key == "lst qrt Brief Risk Decription 3":
                        ws.cell(row=row_num, column=2 + i).value = \
                            risk_combine_text(latest_q_data[name], [
                                'Brief Risk Decription 3', 'BRD 3 Risk Category', 'BRD 3 Primary Risk to',
                                'BRD 3 Internal Control', 'BRD 3 Residual Impact', 'BRD 3 Residual Likelihood'])
                    if key == "lst qrt Brief Risk Decription 4":
                        ws.cell(row=row_num, column=2 + i).value = \
                            risk_combine_text(latest_q_data[name], [
                                'Brief Risk Decription 4', 'BRD 4 Risk Category', 'BRD 4 Primary Risk to',
                                'BRD 4 Internal Control', 'BRD 4 Residual Impact', 'BRD 4 Residual Likelihood'])
                    if key == "lst qrt Brief Risk Decription 5":
                        ws.cell(row=row_num, column=2 + i).value = \
                            risk_combine_text(latest_q_data[name], [
                                'Brief Risk Decription 5', 'BRD 5 Risk Category', 'BRD 5 Primary Risk to',
                                'BRD 5 Internal Control', 'BRD 5 Residual Impact', 'BRD 5 Residual Likelihood'])

                    '''handling of leaders info'''
                    if key == "lst qrt Job Title / Grade":
                        ws.cell(row=row_num, column=2 + i).value = \
                            dn_combine_text(latest_q_data[name], ['Job Title / Grade', 'SRO Grade'])
                    if key == "lst qrt Job Title":
                        ws.cell(row=row_num, column=2 + i).value = \
                            dn_combine_text(latest_q_data[name], ['Job Title', 'PD Grade'])


                    if key == "lst qrt Total Budget/BL":
                        ws.cell(row=row_num, column=2 + i).value = combine_figures(latest_q_data[name],
                                                                                   ['Total Budget/BL', 'Total Forecast'])
                except KeyError:
                    ws.cell(row=row_num, column=40).value = 'Couldnt match'
                ws.cell(row=row_num, column=2 + i).font = red_text
            else:
                ws.cell(row=row_num, column=40).value = 'Couldnt match'

                # this section of the code ensures that all financial costs / benefit forecasts have a zero
                for cost_type in type_list_2:
                    if cost_type in key:
                        try:
                            if latest_q_data[name][key] is None:
                                ws.cell(row=row_num, column=2 + i).value = 0
                        except KeyError:
                            pass

    return workbook

def dn_combine_text(q_data, string_list):

    '''essentaly used to combine two text strings'''

    combined_string = str(q_data[string_list[0]]) + ' - ' + str(q_data[string_list[1]])

    if 'None' in combined_string:
        combined_string = ''

    return combined_string

def strategic_combine_text(q_data, string_list):

    '''essentially used to combined four text strings'''

    combined_string = str(q_data[string_list[0]]) + ' - ' + str(q_data[string_list[1]]) + ' - ' + \
                      str(q_data[string_list[2]]) + ' - ' + str(q_data[string_list[3]])

    if 'None' in combined_string:
        combined_string = ''

    return combined_string

def risk_combine_text(q_data, string_list):

    '''essentially used to combined six text strings'''

    combined_string = str(q_data[string_list[0]]) + ' - ' + str(q_data[string_list[1]]) + \
                      ' - ' + str(q_data[string_list[2]]) + ' - ' + str(q_data[string_list[3]]) + \
                      ' - ' + str(q_data[string_list[4]]) + ' - ' + str(q_data[string_list[5]])

    if 'None' in combined_string:
        combined_string = ''

    return combined_string

def combine_figures(q_data, string_list):

    combined_string = '(B) £' + str(q_data[string_list[0]]) + 'm / (F) £' + str(q_data[string_list[1]])

    return combined_string

latest_dm = load_workbook("C:\\Users\\Standalone\\general\\masters folder\\commission\\Q2_1920_commission_"
                          "master_testing.xlsx")


run = create_master(latest_dm, q1_1920)

run.save("C:\\Users\\Standalone\\general\\masters folder\\commission\\Q2_1920_commission_"
                          "data_testing.xlsx")
