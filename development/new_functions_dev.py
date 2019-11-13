from analysis.data import all_project_names, list_of_masters_all


def get_value(master, project_name, value):
    return master.data[project_name][value]

def baseline_information(project_list, masters_list):

    output = {}

    for project_name in project_list:
        lower_list = []
        for i, master in enumerate(masters_list):
            if project_name in master.projects:
                approved_bc = get_value(master, project_name, 'BICC approval point')
                quarter = get_value(master, project_name, 'Reporting period (GMPP - Snapshot Date)')
                try:
                    previous_approved_bc = get_value(masters_list[i+1], project_name, 'BICC approval point')
                    if approved_bc != previous_approved_bc:
                        lower_list.append((approved_bc, quarter, i))
                except IndexError:
                    # this captures the last available quarter data in the last. For example if a project has been re
                    #
                    lower_list.append((approved_bc, quarter, i))
                except KeyError:
                    lower_list.append((approved_bc, quarter, i))#

        output[project_name] = lower_list

    return output


baseline = baseline_information(all_project_names, list_of_masters_all)
