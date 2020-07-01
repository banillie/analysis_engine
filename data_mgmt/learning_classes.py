from datamaps.api import project_data_from_master
from analysis.engine_functions import baseline_information_bc, baseline_index
from analysis.data import crossrail
from data_mgmt.data import MilestoneData

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


c = crossrail
m = MilestoneData(master_list, bc_index)


