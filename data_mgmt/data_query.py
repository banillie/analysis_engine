'''Returns project values across multiple masters for specified keys of interest. Return for each key is provided
on a separate wb

There are two outputs.
1) wb containing all values
2) wb containing bl values only'''

#TODO color when cell values have changed. refine conditional formatting. bl output specify which qrt bl data taken from

from openpyxl import Workbook
from analysis.data import list_of_masters_all, latest_quarter_project_names, root_path, conditional_text, \
    text_colours, background_colours, baseline_bc_stamp
from analysis.engine_functions import all_milestone_data_bulk, conditional_formatting, get_quarter_stamp

def return_data(project_name_list, data_key_list):
    """Returns project values across multiple masters for specified keys of interest:
    project_names_list: list of project names
    data_key_list: list of data keys
    """
    wb = Workbook()

    for i, key in enumerate(data_key_list):
        '''worksheet is created for each project'''
        ws = wb.create_sheet(key, i)  # creating worksheets
        ws.title = key  # title of worksheet

        '''list project names, groups and stage in ws'''
        for y, project_name in enumerate(project_name_list):

            # get project group info
            try:
                group = list_of_masters_all[0].data[project_name]['DfT Group']
            except KeyError:
                for m, master in enumerate(list_of_masters_all):
                    if project_name in master.projects:
                        group = list_of_masters_all[m].data[project_name]['DfT Group']

            ws.cell(row=2 + y, column=1, value=group) # group info return
            ws.cell(row=2 + y, column=2, value=project_name)  # project name returned

            for x, master in enumerate(list_of_masters_all):
                try:
                    #standard keys
                    if key in list_of_masters_all[x].data[project_name].keys():
                        value = list_of_masters_all[x].data[project_name][key]
                        ws.cell(row=2 + y, column=2 + x, value=value) #retuns value
                        if value is None:
                            ws.cell(row=2 + y, column=2 + x, value='missing data')

                    # milestone keys
                    else:
                        milestones = all_milestone_data_bulk([project_name], list_of_masters_all[x])
                        value = tuple(milestones[project_name][key])[x]
                        ws.cell(row=2 + y, column=2 + x, value=value)
                        ws.cell(row=2 + y, column=2 + x).number_format = 'dd/mm/yy'
                        if value is None:
                            ws.cell(row=2 + y, column=2 + x, value='missing data')

                except KeyError:
                    if project_name in list_of_masters_all[x].projects:
                        #loop calculates if project was not reporting or data missing
                        ws.cell(row=2 + y, column=2 + x, value='missing data')
                    else:
                        ws.cell(row=2 + y, column=2 + x, value='project not reporting')

        '''quarter tag information'''
        ws.cell(row=1, column=1, value='Group')
        ws.cell(row=1, column=2, value='Projects')
        quarter_labels = get_quarter_stamp(list_of_masters_all)
        for l, label in enumerate(quarter_labels):
            ws.cell(row=1, column=l + 3, value=label)

        conditional_formatting(ws, list_columns, conditional_text, text_colours, background_colours, '1', '60')

    return wb

def return_baseline_data(project_name_list, data_key_list):
    wb = Workbook()

    for i, key in enumerate(data_key_list):
        '''worksheet is created for each project'''
        ws = wb.create_sheet(key, i)  # creating worksheets
        ws.title = key  # title of worksheet

        '''list project names, groups and stage in ws'''
        for y, project_name in enumerate(project_name_list):

            # get project group info
            try:
                group = list_of_masters_all[0].data[project_name]['DfT Group']
            except KeyError:
                for m, master in enumerate(list_of_masters_all):
                    if project_name in master.projects:
                        group = list_of_masters_all[m].data[project_name]['DfT Group']

            ws.cell(row=y + 2, column=1, value= group)
            ws.cell(row=y + 2, column=2, value=project_name)  # project name returned


            if key in list_of_masters_all[0].data[project_name].keys():
                # standard keys
                ws.cell(row=y + 2, column=3).value = list_of_masters_all[0].data[project_name][
                    key]  # returns latest value
                for x in range(0, len(baseline_bc_stamp[project_name])):
                    index = baseline_bc_stamp[project_name][x][2]
                    try:
                        ws.cell(row=y + 2, column=x + 4,
                                value=list_of_masters_all[index].data[project_name][key])  # returns baselines
                    except KeyError:
                        ws.cell(row=2 + y, column=4 + x).value = 'missing data'

            else:
                # milestones keys
                milestones = all_milestone_data_bulk([project_name], list_of_masters_all[0])
                try:
                    ws.cell(row=2 + y, column=3).value = tuple(milestones[project_name][key])[0]  # returns latest value
                except KeyError:
                    ws.cell(row=2 + y, column=3).value = 'missing data'

                for x in range(0, len(baseline_bc_stamp[project_name])):
                    index = baseline_bc_stamp[project_name][x][2]
                    try:
                        milestones = all_milestone_data_bulk([project_name], list_of_masters_all[index])
                        ws.cell(row=2 + y, column=3 + x).value = tuple(milestones[project_name][key])[
                            0]  # returns baselines
                    except KeyError:
                        ws.cell(row=2 + y, column=3 + x).value = 'project not reporting'

        ws.cell(row=1, column=1, value='Group')
        ws.cell(row=1, column=2, value='Project')
        ws.cell(row=1, column=3, value='Latest')
        ws.cell(row=1, column=4, value='BL 1')
        ws.cell(row=1, column=5, value='BL 2')
        ws.cell(row=1, column=6, value='BL 3')
        ws.cell(row=1, column=7, value='BL 4')

        conditional_formatting(ws, list_columns, conditional_text, text_colours, background_colours, '1', '60')

    return wb


list_columns = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'q', 's', 't', 'u', 'w']

'''data keys of interest'''
data_interest = ['VfM Category single entry', 'VfM Category lower range', 'VfM Category upper range']

'''Running the programme'''
'''output one - all data'''
run_standard = return_data(latest_quarter_project_names, data_interest)
'''output two - bl data'''
run_baseline = return_baseline_data(latest_quarter_project_names, data_interest)


'''Specify name of the output document here'''
run_standard.save(root_path/'output/data_query_testing.xlsx')
run_baseline.save(root_path/'output/data_query_testing_bl.xlsx')

'''old lists stored here for use in future'''
old_entries = ['GMPP - IPA DCA', 'BICC approval point',
               'Brief project description (GMPP - brief descripton)',
               'Delivery Narrative']

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
