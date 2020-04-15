'''
Returns list of project values for a single master for specified keys of interest. Return is contained within one wb.
Code can handle both standard and project milestone keys, as well as project name lists across
multiple quarters.

There is one outputs.
1) wb containing all values.

Conditional formatting is placed in the files as follows:
rag_rating colours
missing data (md) = black grey
project not reporting (pnr) = light grey
key not collected (knc) = light blue grey
'''


from openpyxl import Workbook
from analysis.data import list_of_masters_all, root_path, gen_txt_list, \
    gen_txt_colours, gen_fill_colours, list_column_ltrs, list_of_rag_keys, rag_txt_list_full, \
    rag_fill_colours, rag_txt_colours, salmon_fill
from analysis.engine_functions import all_milestone_data_bulk, conditional_formatting

def return_data(data_key_list):
    '''
    returns values of interest across multiple ws for baseline values only.
    project_name_list: list of project names
    data_key_list: list of data keys containing values of interest.
    '''
    wb = Workbook()
    ws = wb.active

    '''list project names, groups and stage in ws'''
    for y, project_name in enumerate(list_of_masters_all[0].projects):

        group = list_of_masters_all[0].data[project_name]['DfT Group']

        ws.cell(row=2 + y, column=1, value=group) # group info
        ws.cell(row=2 + y, column=2, value=project_name)  # project name returned

        for x, key in enumerate(data_key_list):
            ws.cell(row=1, column=3 + x, value=key)
            try: # standard keys
                value = list_of_masters_all[0].data[project_name][key]
                if value is None:
                    ws.cell(row=2 + y, column=3 + x).value = 'md'
                else:
                    ws.cell(row=2 + y, column=3 + x, value=value)
                try:  # checks for change against last quarter
                    lst_value = list_of_masters_all[1].data[project_name][key]
                    if value != lst_value:
                        ws.cell(row=2 + y, column=3 + x).fill = salmon_fill
                except (KeyError, IndexError):
                    pass
            except KeyError:
                try: # milestone keys
                    milestones = all_milestone_data_bulk([project_name], list_of_masters_all[0])
                    value = tuple(milestones[project_name][key])[0]
                    if value is None:
                        ws.cell(row=2 + y, column=3 + x).value = 'md'
                    else:
                        ws.cell(row=2 + y, column=3 + x).value = value
                        ws.cell(row=2 + y, column=3 + x).number_format = 'dd/mm/yy'
                    try:  # loop checks if value has changed since last quarter
                        old_milestones = all_milestone_data_bulk([project_name], list_of_masters_all[1])
                        lst_value = tuple(old_milestones[project_name][key])[0]
                        if value != lst_value:
                            ws.cell(row=2 + y, column=3 + x).fill = salmon_fill
                    except (KeyError, IndexError):
                        pass
                except KeyError: # exception catches both standard and milestone keys
                    ws.cell(row=2 + y, column=3 + x).value = 'knc'
                except TypeError:
                    ws.cell(row=2 + y, column=3 + x).value = 'pnr'

    for z, key in enumerate(data_key_list):
        if key in list_of_rag_keys:
            conditional_formatting(ws, [list_column_ltrs[z+2]], rag_txt_list_full, rag_txt_colours, rag_fill_colours,
                                   '1', '60') # plus 2 in column ltrs as values start being placed in at col 2.
    '''quarter tag information'''
    ws.cell(row=1, column=1, value='Group')
    ws.cell(row=1, column=2, value='Projects')

    conditional_formatting(ws, list_column_ltrs, gen_txt_list, gen_txt_colours, gen_fill_colours, '1', '60')

    return wb

'''data keys of interest. Place all keys of interest as stings in this list'''
data_interest = ['IPDC approval point',
                 'Total Forecast',
                 'VfM Category single entry',
                 'VfM Category lower range',
                 'VfM Category upper range',
                 'Present Value Cost (PVC)',
                 'Present Value Benefit (PVB)',
                 'Initial Benefits Cost Ratio (BCR)',
                 'Adjusted Benefits Cost Ratio (BCR)',
                 'Start of Construction/build',
                 'Start of Operation',
                 'Full Operations',
                 'Project End Date']

'''Running the programme'''

'''output one - all data'''
run_standard = return_data(data_interest)

'''Specify name of the output document here. See general guidance re saving output files'''
run_standard.save(root_path/'output/su_data_query.xlsx')

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
