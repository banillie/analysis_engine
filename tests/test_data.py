"""
Tests for analysis_engine
"""

from data_mgmt.data import MilestoneData, MilestoneChartData, Masters, CostData, \
    Projects
import pytest
import datetime
from datamaps.api import project_data_from_master

start_date = datetime.date(2020, 6, 1)
end_date = datetime.date(2022, 6, 30)

def test_creation_of_Masters_class(basic_master, abbreviations):
    projects = list(abbreviations.keys())
    master = Masters(basic_master, projects)
    assert isinstance(master.master_data, (list,))
    assert master.project_names == ['Sea of Tranquility', 'Apollo 11', 'Apollo 13', 'Falcon 9', 'Columbia', 'Mars']

def test_getting_baseline_data_from_Masters(basic_master, abbreviations):
    projects = list(abbreviations.keys())
    master = Masters(basic_master, projects)
    master.baseline_data('Re-baseline IPDC milestones')
    assert isinstance(master.bl_index, (dict,))
    assert master.bl_index['Sea of Tranquility'] == [0, 1]
    assert master.bl_index['Apollo 11'] == [0, 1, 2]


def test_MilestoneData_project_dict_returns_dict(basic_master, abbreviations):
    m = Masters(basic_master, basic_master[0].projects)
    m.baseline_data("Re-baseline IPDC milestones")
    m = MilestoneData(m, abbreviations)
    assert isinstance(m.project_current, (dict,))
#
#
# def test_MilestoneData_group_dict_returns_dict(mst, abbreviations):
#     mst.baseline_data('Re-baseline IPDC milestones')
#     m = MilestoneData(mst, abbreviations)
#     assert isinstance(m.group_current, (dict,))
#
#
# def test_MilestoneChartData_group_chart_returns_list(mst, abbreviations):
#     mst.baseline_data('Re-baseline IPDC milestones')
#     m = MilestoneData(mst, abbreviations)
#     mcd = MilestoneChartData(milestone_data_object=m)
#     assert isinstance(mcd.group_current_tds, (list,))
#
#
# def test_MilestoneChartData_group_chart_filter_in_works(mst, abbreviations):
#     assurance = ['Gateway', 'SGAR', 'Red', 'Review']
#     mst.baseline_data('Re-baseline IPDC milestones')
#     m = MilestoneData(mst, abbreviations)
#     mcd = MilestoneChartData(m, keys_of_interest=assurance)
#     assert any("Gateway" in s for s in mcd.group_keys)
#     assert any("SGAR" in s for s in mcd.group_keys)
#     assert any("Red" in s for s in mcd.group_keys)
#     assert any("Review" in s for s in mcd.group_keys)
#
#
# def test_MilestoneChartData_group_chart_filter_out_works(mst, abbreviations):
#     assurance = ['Gateway', 'SGAR', 'Red', 'Review']
#     mst.baseline_data('Re-baseline IPDC milestones')
#     m = MilestoneData(mst, abbreviations)
#     mcd = MilestoneChartData(m, keys_not_of_interest=assurance)
#     assert not any("Gateway" in s for s in mcd.group_keys)
#     assert not any("SGAR" in s for s in mcd.group_keys)
#     assert not any("Red" in s for s in mcd.group_keys)
#     assert not any("Review" in s for s in mcd.group_keys)
#
#
# def test_CostData_cost_total_spent_returns_lists(mst):
#     mst.baseline_data('Re-baseline IPDC cost')
#     c = CostData(mst)
#     assert isinstance(c.spent, (list,))
#     assert isinstance(c.cat_spent, (list,))
#
#
# def test_ProjectsGroupName_returns_a12():
#     assert Projects.a12 == 'A12 Chelmsford to A120 widening'
#
#
# def test_ProjectGroupName_returns_rpe_as_list():
#     assert isinstance(Projects.rpe, (list,))
