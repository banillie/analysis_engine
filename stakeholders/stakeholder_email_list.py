'''
programme for getting email list of stakeholders from the portfolio contact list spreadsheet
'''

import re
from analysis.data import q2_1920, all_project_names
from analysis.engine_functions import filter_project_group

def get_stakeholder_list(stakeholder_dict, project_list):
    output_list = []

    for project in project_list:
        print(project)
        commission = stakeholder_dict[project]['Commission']
        match_com = re.findall(r'[\w\.-]+@[\w\.-]+', commission)
        for email in match_com:
            output_list.append(email)

        cc = stakeholder_dict[project]['Commission CC']
        match_cc = re.findall(r'[\w\.-]+@[\w\.-]+', cc)
        for email in match_cc:
            output_list.append(email)

        poc = stakeholder_dict[project]['Working Contact email']
        print(poc)
        match_poc = re.findall(r'[\w\.-]+@[\w\.-]+', poc)
        for email in match_poc:
            output_list.append(email)

        pd = stakeholder_dict[project]['PD email']
        match_pd = re.findall(r'[\w\.-]+@[\w\.-]+', pd)
        for email in match_pd:
            output_list.append(email)

        sro = stakeholder_dict[project]['SRO email']
        match_sro = re.findall(r'[\w\.-]+@[\w\.-]+', sro)
        for email in match_sro:
            output_list.append(email)

    final_list = list(set(output_list))
    final_list_one = sorted(final_list)
    final_final_list = sorted(final_list_one, key=str.lower)

    return final_final_list

def filter_mode(stakeholder_dict, mode):
    output_list = []

    for project in stakeholder_dict.keys():
        if stakeholder_dict[project]['Mode'] == mode:
            output_list.append(project)

    return output_list


email_list = get_stakeholder_list(stakeholders, all_projects)