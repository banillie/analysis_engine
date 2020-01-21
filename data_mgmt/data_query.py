'''

Programme for querying and returning non-milestone data from master data set.

There are several options that need to be specified below:
1) returning single data of interest (in to one tab)
2) returning several data of interest (across multiple tabs)... in development 
3) return data across all masters 
4) return data all pertaining to latest, last and baseline data. 

some formatting is placed into the output file:
2) changes in reported data - highlighted by salmon pink background,
3) when projects were not reporting data -  grey out cell,
4) if a rag status is returned the colour of the rag status

Follow instruction as set out below are provided


'''

from openpyxl import Workbook
from analysis.data import list_of_masters_all, latest_quarter_project_names, bc_index, baseline_bc_stamp, salmon_fill, \
    root_path
from analysis.engine_functions import get_all_project_names, get_quarter_stamp, grey_conditional_formatting

def return_data(project_name_list, data_key_list):
    '''
    places all (non-milestone) data of interest into excel file output
    project_name_list: list of projects to return data for
    data_key_list: the data key of interest

    '''

    wb = Workbook()

    '''project data into ws'''
    for i, data_key in enumerate(data_key_list):
        ws = wb.create_sheet(data_key, i)  # creating worksheets
        ws.title = data_key  # title of worksheet

        '''lists project names in ws'''
        for x in range(0, len(project_name_list)):
            ws.cell(row=x + 2, column=2, value=project_name_list[x])
            try:
                ws.cell(row=x + 2, column=1).value = list_of_masters_all[0].data[project_name_list[x]]['DfT Group']
            except KeyError:
                pass


        for row_num in range(2, ws.max_row + 1):
            project_name = ws.cell(row=row_num, column=2).value
            print(project_name)
            col_start = 3
            for i, master in enumerate(list_of_masters_all):
                if project_name in master.projects:
                    try:
                        ws.cell(row=row_num, column=col_start).value = master.data[project_name][data_key]
                        if master.data[project_name][data_key] is None:
                            ws.cell(row=row_num, column=col_start).value = 'None'
                        try:
                            if list_of_masters_all[i+1].data[project_name][data_key] != master.data[project_name][data_key]:
                                ws.cell(row=row_num, column=col_start).fill = salmon_fill
                        except (IndexError, KeyError):
                            pass
                        col_start += 1
                    except KeyError:
                        ws.cell(row=row_num, column=col_start).value = 'Data not collected'
                        col_start += 1
                else:
                    ws.cell(row=row_num, column=col_start).value = 'Not reporting'
                    col_start += 1

        '''quarter tag / meta data into ws'''
        quarter_labels = get_quarter_stamp(list_of_masters_all)
        ws.cell(row=1, column=1, value='Group')
        ws.cell(row=1, column=2, value='Project')
        for i, label in enumerate(quarter_labels):
            ws.cell(row=1, column=i + 3, value=label)

        grey_conditional_formatting(ws)  # apply grey formatting

    return wb

def return_baseline_data(project_name_list, data_key_list):
    '''
    places all non-milestone data into output document with latest, last and baseline data. Also states which quarter
    is being used as baseline
    :param project_name_list: list of project names
    :param data_key_list: data of interest/to return
    :return: excel spreadsheet
    '''

    wb = Workbook()

    '''project data into ws'''
    for i, data_key in enumerate(data_key_list):
        ws = wb.create_sheet(data_key, i)  # creating worksheets
        ws.title = data_key  # title of worksheet

        '''lists project names in ws'''
        for x in range(0, len(project_name_list)):
            ws.cell(row=x + 2, column=2, value=project_name_list[x])
            try:
                ws.cell(row=x + 2, column=1).value = list_of_masters_all[0].data[project_name_list[x]]['DfT Group']
            except KeyError:
                pass

        '''project data into ws'''
        for row_num in range(2, ws.max_row + 1):
            project_name = ws.cell(row=row_num, column=2).value
            # ref to baseline quarter
            ws.cell(row=row_num, column=8).value = baseline_bc_stamp[project_name][0][1]
            print(project_name)
            col_start = 3
            for i in bc_index[project_name]:
                try:
                    ws.cell(row=row_num, column=col_start).value = list_of_masters_all[i].data[project_name][data_key]
                    if list_of_masters_all[i].data[project_name][data_key] is None:
                        ws.cell(row=row_num, column=col_start).value = 'None'
                    try:
                        if list_of_masters_all[i+1].data[project_name][data_key] != list_of_masters_all[i].data[project_name][data_key]:
                            ws.cell(row=row_num, column=col_start).fill = salmon_fill
                    except (IndexError, KeyError):
                        pass
                    col_start += 1
                except (KeyError, TypeError):
                    ws.cell(row=row_num, column=col_start).value = 'Data not collected'
                    col_start += 1

        '''quarter tag / meta data into ws'''
        baseline_labels = ['This quarter', 'Last quarter', 'Baseline quarter']
        ws.cell(row=1, column=1, value='Group')
        ws.cell(row=1, column=2, value='Project')
        for i, label in enumerate(baseline_labels):
            ws.cell(row=1, column=i + 3, value=label)
        ws.cell(row=1, column=8, value='Quarter from which baseline data taken')

    return wb

''' RUNNING PROGRAMME '''

'''Note that the all master data is taken from the data file. Make sure that this is up to date and that all relevant
  data is being imported'''

''' ONE. Set relevant list of projects. This needs to be done in accordance with the data you are working with via the
 data.py file '''
one_quarter_list = latest_quarter_project_names
combined_quarters_list = get_all_project_names(list_of_masters_all)
specific_project_list = [] # list of projects. get project name in import statement

'''TWO. Set data of interest. there are two options here. hash out whichever option you are not using'''
'''option one - non-milestone data. NOTE. this must be in a list [] even if just one data key'''
data_interest = ['Departmental DCA', 'Working Contact Name', 'Working Contact Email',
                 'Brief project description (GMPP - brief descripton)',
                 'Business Case & Version No.', 'BICC approval point',
                 'NPV for all projects and NPV for programmes if available',
                 'Initial Benefits Cost Ratio (BCR)', 'Adjusted Benefits Cost Ratio (BCR)',
                 'VfM Category single entry', 'VfM Category lower range', 'VfM Category upper range',
                 'Present Value Cost (PVC)', 'Present Value Benefit (PVB)', 'SRO Benefits RAG',
                 'Benefits Narrative', 'Ben comparison with last quarters cost - narrative']

'''THREE. Run the programme'''
'''option one - run the return_data function for all non-milestone data'''
run_standard = return_data(one_quarter_list, data_interest)

'''option two - run the return_baseline_data function for all non-milestone data'''
run_baseline = return_baseline_data(one_quarter_list, data_interest)

'''FOUR. specify the file path and name of the output document'''
run_standard.save(root_path/'output/q3_1920_vfm_data.xlsx')

run_baseline.save(root_path/'output/q3_1920_vfm_baseline_data.xlsx')

'''old lists stored here for use in future'''

old_entries = ['GMPP - IPA DCA', 'BICC approval point']

vfm_analysis_list = ['Departmental DCA', 'Working Contact Name', 'Working Contact Email',
                 'Brief project description (GMPP - brief descripton)',
                 'Business Case & Version No.', 'BICC approval point',
                 'NPV for all projects and NPV for programmes if available',
                 'Initial Benefits Cost Ratio (BCR)', 'Adjusted Benefits Cost Ratio (BCR)',
                 'VfM Category single entry', 'VfM Category lower range', 'VfM Category upper range',
                 'Present Value Cost (PVC)', 'Present Value Benefit (PVB)', 'SRO Benefits RAG',
                 'Benefits Narrative', 'Ben comparison with last quarters cost - narrative']

ipa_ar_fields_1920 =  ['Department', '19-20 RDEL BL Total', '19-20 CDEL BL WLC',
                 '19-20 RDEL Forecast Total', '19-20 CDEL Forecast Total WLC', 'Total BL',
                 'GMPP - IPA ID Number']
