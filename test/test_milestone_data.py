from data_mgmt.data import MilestoneData
from datamaps.api import project_data_from_master
from analysis.data import root_path, bc_index


q4_1920 = project_data_from_master(root_path/'core_data/master_4_2019.xlsx', 4, 2019)
project_names = q4_1920.projects
master_data = [q4_1920]


def test_project_names_appear_in_object_project_names_attribute():
    m = MilestoneData(master_data, bc_index, 0)
    assert "A12 Chelmsford to A120 widening" in m.project_names


def test_baseline_index():
    m = MilestoneData(master_data, bc_index, 0)
    assert isinstance(m.baseline_index, (dict,))

def test_get_project_dict():
    m = MilestoneData(master_data, bc_index, 0)
    d = m.get_project_dict()
    assert x == "?"


