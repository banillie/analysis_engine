from datamaps.api import project_data_from_master
from analysis.engine_functions import baseline_information_bc, baseline_index
from analysis.data import crossrail
import platform
from pathlib import Path


'''file path'''
def _platform_docs_dir() -> Path:
    if platform.system() == "Linux":
        return Path.home() / "Documents" / "analysis_engine"
    if platform.system() == "Darwin":
        return Path.home() / "Documents" / "analysis_engine"
    else:
        return Path.home() / "Documents" / "analysis_engine"

root_path = _platform_docs_dir()

q4_1920 = project_data_from_master(root_path/'core_data/master_4_2019.xlsx', 4, 2019)
q3_1920 = project_data_from_master(root_path/'core_data/master_3_2019.xlsx', 3, 2019)
q2_1920 = project_data_from_master(root_path/'core_data/master_2_2019.xlsx', 2, 2019)
q1_1920 = project_data_from_master(root_path/'core_data/master_1_2019.xlsx', 1, 2019)
q4_1819 = project_data_from_master(root_path/'core_data/master_4_2018.xlsx', 4, 2018)

master_list = [q4_1920,
               q3_1920,
               q2_1920,
               q1_1920,
               q4_1819]

p_names = q4_1920.projects
# general baseline information
baseline_bc_stamp = baseline_information_bc(p_names, master_list)
bc_index = baseline_index(baseline_bc_stamp, master_list)

class Data:
    def __init__(self, master_data=list):
        self.master_data = master_data
        # self.quarter_data = ''
        # self.get_quarter_data()

    # def get_quarter_data(self, quarter=str):
    #     for i in len(self.master_data):
    #         if quarter == str(self.master_data[i].quarter):
    #             self.quarter_data = self.master_data[i]
    #
    #     return self.quarter_data

    # def get_project_names(self,):
    #
    # def get_bc_approvals(self, ):


class MilestoneData:
    def __init__(self, master_data):
        self.master_data = master_data

    def get_project_dict(project_names, baseline_index, data_to_return):
        self.project_dict = {}
        self.get_project_dict()
        # self.group_dict = {}

        upper_dict = {}

        for name in self.project_names:
            lower_dict = {}
            raw_list = []
            try:
                p_data = self.master_data[self.baseline_index[name][self.data_to_return]].data[name]
                for i in range(1, 50):
                    try:
                        try:
                            t = (p_data['Approval MM' + str(i)],
                                 p_data['Approval MM' + str(i) + ' Forecast / Actual'],
                                 p_data['Approval MM' + str(i) + ' Notes'])
                            raw_list.append(t)
                        except KeyError:
                            t = (p_data['Approval MM' + str(i)],
                                 p_data['Approval MM' + str(i) + ' Forecast - Actual'],
                                 p_data['Approval MM' + str(i) + ' Notes'])
                            raw_list.append(t)

                        t = (p_data['Assurance MM' + str(i)],
                             p_data['Assurance MM' + str(i) + ' Forecast - Actual'],
                             p_data['Assurance MM' + str(i) + ' Notes'])
                        raw_list.append(t)

                    except KeyError:
                        pass

                for i in range(18, 67):
                    try:
                        t = (p_data['Project MM' + str(i)],
                             p_data['Project MM' + str(i) + ' Forecast - Actual'],
                             p_data['Project MM' + str(i) + ' Notes'])
                        raw_list.append(t)
                    except KeyError:
                        pass
            except (KeyError, TypeError):
                print('yes')
                pass

            # put the list in chronological order
            sorted_list = sorted(raw_list, key=lambda k: (k[1] is None, k[1]))
            print(sorted_list)

            # loop to stop key names being the same. Not ideal as doesn't handle keys that may already have numbers as
            # strings at end of names. But still useful.
            for x in sorted_list:
                if x[0] is not None:
                    if x[0] in lower_dict:
                        for i in range(2, 15):
                            key_name = x[0] + ' ' + str(i)
                            if key_name in lower_dict:
                                continue
                            else:
                                lower_dict[key_name] = {x[1]: x[2]}
                                break
                    else:
                        lower_dict[x[0]] = {x[1]: x[2]}
                else:
                    pass

            upper_dict[name] = lower_dict

        self.project_dict = upper_dict

        return self.project_dict

    # def get_data_for_project(self, project_name):
    #     return self.dict[project_name]

c = crossrail
#d = Data(master_list)
m = MilestoneData(p_names, bc_index, 0)
#m.get_data_for_project("A12 Extension")