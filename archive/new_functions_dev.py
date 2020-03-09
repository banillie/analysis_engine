'''archiving... can probably delete. think all code has been intergrated into analysis_engine as necessary'''

from analysis.data import latest_quarter_projects, list_of_masters_all, latest_cost_profiles, cost_list, year_list


def baseline_information_bc(project_list, masters_list):
    '''
    Function that calculates when project business case of have changes. Only returns where there have been changes.
    :param project_list: list of project names
    :param masters_list: list of masters with quarter information
    :return: python dictionary in format 'project name':('BC', 'Quarter Stamp', index position in masters_list)
    '''
    output = {}

    for project_name in project_list:
        lower_list = []
        for i, master in enumerate(masters_list):
            if project_name in master.projects:
                approved_bc = master.data[project_name]['BICC approval point']
                quarter = master.data[project_name]['Reporting period (GMPP - Snapshot Date)']
                try:
                    previous_approved_bc = masters_list[i+1].data[project_name]['BICC approval point']
                    if approved_bc != previous_approved_bc:
                        lower_list.append((approved_bc, quarter, i))
                except IndexError:
                    # this captures the last available quarter data if project was in portfolio from beginning
                    lower_list.append((approved_bc, quarter, i))
                except KeyError:
                    # this captures the first quarter the project reported if not in portfolio from beginning
                    lower_list.append((approved_bc, quarter, i))

        output[project_name] = lower_list

    return output

def baseline_information(project_list, masters_list, data_baseline):
    '''
    Function that calculates in information within masters has been baselined
    :param project_list: list of project names
    :param masters_list: list of quarter masters.
    :param data_baseline: type of information to check for baselines. options are: 'ALB milestones' etc
    :return: python dictionary structured as 'project name': ('yes (if reported that quarter), 'quarter stamp',
    index position of master in master quarter list)
    '''

    output = {}

    for project_name in project_list:
        lower_list = []
        for i, master in enumerate(masters_list):
            if project_name in master.projects:
                try:
                    approved_bc = master.data[project_name]['Re-baseline ' + data_baseline]
                    quarter = master.data[project_name]['Reporting period (GMPP - Snapshot Date)']
                    if approved_bc == 'Yes':
                        lower_list.append((approved_bc, quarter, i))
                except KeyError:
                    pass

        output[project_name] = lower_list

    return output


def baseline_index(baseline_data):
    '''
    Function that calculates the index list for baseline data
    :param baseline_data: output created by either baseline_information or baseline_information_bc functions
    :return: python dictionary in format 'project name':[index list]
    '''

    output = {}

    for project_name in baseline_data:
        lower_list = [0, 1]
        for tuple_info in baseline_data[project_name]:
            lower_list.append(tuple_info[2])

        output[project_name] = lower_list

    return output


def calculate_group_project_total(project_name_list, master_data, project_name_no_count_list, type_list):
    '''
    calculates the total cost figure for each year and type of spend e.g. RDEL 19-20, for all projects of interest.
    :param project_name_list: list of project names
    :param master_data: master data set as created by the get_project_cost_profile
    :param project_name_no_count_list: list of project names to remove from total figures, to ensure no double counting
    e.g. if there are separate schemes as well as overall programme reporting.
    :param type_list: the type of financial figure list being counted. e.g. costs or income
    :return: python dictionary in format 'year + spend type': total
    '''

    output = {}

    project_list = [x for x in project_name_list if x not in project_name_no_count_list]

    for cost in type_list:
        for year in year_list:
            total = 0
            for project_name in project_list:
                total = total + master_data[project_name][year + cost]


            output[ year + cost] = total

    return output

dont_double_count = ['HS2 Phase 2b', 'HS2 Phase1', 'HS2 Phase2a', 'East Midlands Franchise',
                     'West Coast Partnership Franchise', 'Northern Powerhouse Rail', 'East Coast Digital Programme']


#baseline = baseline_information_bc(all_project_names, list_of_masters_all)

#index = baseline_index(baseline)

test = calculate_group_project_total(latest_quarter_projects, latest_cost_profiles, dont_double_count, cost_list)