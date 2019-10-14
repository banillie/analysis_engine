'''programme for altering across all masters changes in milestone keys.

working. bit of a hack at the mo though. further work required to finalise it and make pretty!

'''

#TODO get code to handle none data.

from bcompiler.utils import project_data_from_master
from openpyxl import load_workbook
from openpyxl.styles import Font

'''Function to filter out ALL milestone data'''
def all_milestone_data_bulk(project_list, master_data):
    upper_dict = {}

    for name in project_list:
        try:
            p_data = master_data[name]
            lower_dict = {}
            for i in range(1, 50):
                try:
                    try:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast / Actual']: p_data[
                                'Approval MM' + str(i) + ' Notes']}
                    except KeyError:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast - Actual']: p_data[
                                'Approval MM' + str(i) + ' Notes']}

                    lower_dict[p_data['Assurance MM' + str(i)]] = \
                        {p_data['Assurance MM' + str(i) + ' Forecast - Actual']: p_data[
                                'Assurance MM' + str(i) + ' Notes']}
                except KeyError:
                    pass

            for i in range(18, 67):
                try:
                    lower_dict[p_data['Project MM' + str(i)]] = \
                        {p_data['Project MM' + str(i) + ' Forecast - Actual']: p_data['Project MM' + str(i) + ' Notes']}
                except KeyError:
                    pass
        except KeyError:
            lower_dict = {}

        upper_dict[name] = lower_dict

    return upper_dict

'''Function to filter out approval and project delivery milestones'''
def ap_p_milestone_data_bulk(project_list, master_data):
    upper_dict = {}

    for name in project_list:
        try:
            p_data = master_data[name]
            lower_dict = {}
            for i in range(1, 50):
                try:
                    try:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast / Actual'] : p_data['Approval MM' + str(i) + ' Notes']}
                    except KeyError:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast - Actual'] : p_data['Approval MM' + str(i) + ' Notes']}

                except KeyError:
                    pass

            for i in range(18, 67):
                try:
                    lower_dict[p_data['Project MM' + str(i)]] = \
                        {p_data['Project MM' + str(i) + ' Forecast - Actual'] : p_data['Project MM' + str(i) + ' Notes']}
                except KeyError:
                    pass
        except KeyError:
            lower_dict = {}

        upper_dict[name] = lower_dict

    return upper_dict

'''Function to filter out assurance milestone data'''
def assurance_milestone_data_bulk(project_list, master_data):
    upper_dict = {}

    for name in project_list:
        try:
            p_data = master_data[name]
            lower_dict = {}
            for i in range(1, 50):
                lower_dict[p_data['Assurance MM' + str(i)]] = \
                    {p_data['Assurance MM' + str(i) + ' Forecast - Actual']: p_data['Assurance MM' + str(i) + ' Notes']}

            upper_dict[name] = lower_dict
        except KeyError:
            upper_dict[name] = {}

    return upper_dict

'''Function that calculates time different between milestone dates'''
def project_time_difference(proj_m_data_1, proj_m_data_2, date_of_interest):
    upper_dict = {}

    for proj_name in proj_m_data_1:
        td_dict = {}
        for milestone in proj_m_data_1[proj_name]:
            if milestone is not None:
                milestone_date = tuple(proj_m_data_1[proj_name][milestone])[0]
                try:
                    if date_of_interest <= milestone_date:
                        try:
                            old_milestone_date = tuple(proj_m_data_2[proj_name][milestone])[0]
                            time_delta = (milestone_date - old_milestone_date).days  # time_delta calculated here
                            if time_delta == 0:
                                td_dict[milestone] = 0
                            else:
                                td_dict[milestone] = time_delta
                        except (KeyError, TypeError):
                            td_dict[milestone] = 'Not reported' # not reported that quarter
                except (KeyError, TypeError):
                    td_dict[milestone] = 'No date provided' # date has now been removed

        upper_dict[proj_name] = td_dict

    return upper_dict


''' One of key functions used for calculating which quarter to baseline data from...
Function returns a dictionary structured in the following way project name[('latest quarter info', 'latest bc'), 
('last quarter info', 'last bc'), ('last baseline quarter info', 'last baseline bc'), ('oldest quarter info', 
'oldest bc')] depending on the amount information available in the data. Only the first three key values are returned, 
to ensure consistency (which is helpful later).'''
def bc_ref_stages(proj_list, q_masters_dict_list):

    output_dict = {}

    for name in proj_list:
        #print(name)
        all_list = []      # format [('quarter info': 'bc')] across all masters including project
        bl_list = []        # format ['bc', 'bc'] across all masters. bl_list_2 removes duplicates
        ref_list = []       # format as for all list but only contains the three tuples of interest
        for master in q_masters_dict_list:
            try:
                bc_stage = master[name]['BICC approval point']
                quarter = master[name]['Reporting period (GMPP - Snapshot Date)']
                tuple = (quarter, bc_stage)
                all_list.append(tuple)
            except KeyError:
                pass

        for i in range(0, len(all_list)):
            bl_list.append(all_list[i][1])

        '''below lines of text from stackoverflow. Question, remove duplicates in python list while 
        preserving order'''
        seen = set()
        seen_add = seen.add
        bl_list_2 = [x for x in bl_list if not (x in seen or seen_add(x))]

        ref_list.insert(0, all_list[0])     # puts the latest info into the list first

        try:
            ref_list.insert(1, all_list[1])    # puts that last info into the list
        except IndexError:
            ref_list.insert(1, all_list[0])

        if len(bl_list_2) == 1:                     # puts oldest info into list (as basline if no baseline)
            ref_list.insert(2, all_list[-1])
        else:
            for i in range(0, len(all_list)):      # puts in baseline
                if all_list[i][1] == bl_list[0]:
                    ref_list.insert(2, all_list[i])

        '''there is a hack here i.e. returning only first three in ref_list. There's a bug which I don't fully 
        understand, but this solution is hopefully good enough for now'''
        output_dict[name] = ref_list[0:3]

    return output_dict

'''Another key function used for calcualting which quarter to baseline data from...
Fuction returns a dictionay structured in the following way project_name[n,n,n]. The n (number) values denote where 
the relevant quarter master dictionary is positions in the list of master dictionaries'''
def get_master_baseline_dict(proj_list, q_masters_dict_list, baseline_dict_list):
    output_dict = {}

    for name in proj_list:
        master_q_list = []
        for key in baseline_dict_list[name]:
            for x, master in enumerate(q_masters_dict_list):
                try:
                    quarter = master[name]['Reporting period (GMPP - Snapshot Date)']
                    if quarter == key[0]:
                        master_q_list.append(x)
                except KeyError:
                    pass

        output_dict[name] = master_q_list

    return output_dict

def change_key(proj_list, q_master_wb_title_list, baseline_list, key_change_dict):
    red_text = Font(color="00fc2525")
    for proj_name in proj_list:
        for index in baseline_list[proj_name]:
            print(index)
            wb = load_workbook(q_master_wb_title_list[index])
            ws = wb.active
            for col_num in range(2, ws.max_column + 1):
                project_name = ws.cell(row=1, column=col_num).value
                #print(project_name)
                if project_name == proj_name:
                    print(proj_name)
                    for row_num in range(2, ws.max_row + 1):
                        for i in range(1, 6):
                            try:
                                if ws.cell(row=row_num, column=col_num).value == key_change_dict[proj_name]['Key '+ str(i)]:
                                    print(key_change_dict[proj_name]['Key '+ str(i)])
                                    ws.cell(row=row_num, column=col_num).value = key_change_dict[proj_name]['Key '+ str(i)+' change']
                                    print(key_change_dict[proj_name]['Key '+ str(i)+' change'])
                                    ws.cell(row=row_num, column=col_num).font = red_text
                                else:
                                    pass
                            except KeyError:
                                pass

            wb.save(q_master_wb_title_list[index])


'''INSTRUCTIONS FOR RUNNING THE PROGRAMME'''

'''1) load all master quarter data files here. They have to be store twice. The first time they are converted into
dictionaries. The second time the filepath is stored as a part of the list (this is so the current masters can be 
opened amended and saved again, make sure lists are identical'''

'''i) dictionaries'''
q1_1920 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2019_wip'
                                   '_(25_7_19).xlsx')
q4_1819 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2018.xlsx')
q3_1819 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2018.xlsx')
q2_1819 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_2_2018.xlsx')
q1_1819 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2018.xlsx')
q4_1718 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2017.xlsx')
q3_1718 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2017.xlsx')
q2_1718 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_2_2017.xlsx')
q1_1718 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2017.xlsx')
q4_1617 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2016.xlsx')
q3_1617 = project_data_from_master('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2016.xlsx')

'''ii) list of file paths to masters'''
master_list = ('C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2019_wip_(25_7_19).xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2018.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2018.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_2_2018.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2018.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2017.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2017.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_2_2017.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_1_2017.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_4_2016.xlsx',
               'C:\\Users\\Standalone\\general\\masters folder\\core data\\master_3_2016.xlsx')

'''2) Put the masters dictionaries into a list. '''
list_of_dicts_all = [q1_1920, q4_1819, q3_1819, q2_1819, q1_1819, q4_1718, q3_1718, q2_1718, q1_1718, q4_1617, q3_1617]
#list_of_dicts_bespoke = [zero, last]

'''3) provide file path to document which contains information on the data that needs to be changed'''
key_change = project_data_from_master('C:\\Users\\Standalone\\general\\change_milestone_key_testing.xlsx')

'''4) list of projects. taken from the key change document - as this contains the only projects that need information 
changed'''
proj_list = list(key_change.keys())
proj_list_bespoke = ['South West Route Capacity', 'North of England Programme']

'''ignore this part. no change required'''
baseline_bc = bc_ref_stages(proj_list, list_of_dicts_all)
q_master_baseline_no = get_master_baseline_dict(proj_list, list_of_dicts_all, baseline_bc)

'''5) enter relevant variables in the change_key function'''
change_key(proj_list, master_list, q_master_baseline_no, key_change)