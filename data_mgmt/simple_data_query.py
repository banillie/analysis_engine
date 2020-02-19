'''development code for returning multiple data key values (for one quarter) into one wb'''

# TODO. change date format in excel wb

from openpyxl import Workbook
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from analysis.data import list_of_masters_all, latest_quarter_project_names, bc_index, baseline_bc_stamp, salmon_fill, \
    root_path
from analysis.engine_functions import all_milestone_data_bulk, ap_p_milestone_data_bulk, assurance_milestone_data_bulk,\
    get_all_project_names, get_quarter_stamp, grey_conditional_formatting

def return_data(project_name_list, data_key_list):

    wb = Workbook()
    ws = wb.active

    for i, project_name in enumerate(project_name_list):
        '''list project names, groups and stage in ws'''
        ws.cell(row=2 + i, column=1).value = project_name
        for y, key in enumerate(data_key_list):
            ws.cell(row=1, column=2+y, value=key) #returns key

            try:
                #standard keys
                if key in list_of_masters_all[0].data[project_name].keys():
                    value = list_of_masters_all[0].data[project_name][key]
                    ws.cell(row=2+i, column=2+y, value=value) #retuns value

                # milestone keys
                else:
                    milestones = all_milestone_data_bulk([project_name], list_of_masters_all[0])
                    value = tuple(milestones[project_name][key])[0]
                    print(value)
                    ws.cell(row=2 + i, column=2 + y, value=value)
                    ws.cell(row=2 + i, column=2 + y).number_format = 'dd/mm/yy'

            except KeyError:
                if project_name in list_of_masters_all[0].projects: #loop calculates if project was not reporting or data missing
                    ws.cell(row=2+ i, column=2+y, value='missing data')
                else:
                    ws.cell(row=2+i, column=2+y, value='project not reporting')

        '''quarter tag information'''
        ws.cell(row=1, column=1, value='Project')


    return wb


'''data keys of interest'''
key_list = ['BICC approval point',
            'SRO Full Name',
            'SRO Email',
            'SRO Phone No.',
            'Brief project description (GMPP - brief descripton)',
            'Project End Date']

'''running prog - step one'''
run_project_all = return_data(latest_quarter_project_names, key_list)



'''step two'''
run_project_all.save(root_path/'output/ties_data.xlsx')
