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

# test_master_one = project_data_from_master("/home/will/code/python/analysis_engine/tests/resources/test_master_4_2016.xlsx", 4, 2016)
# test_master_two = project_data_from_master("/home/will/code/python/analysis_engine/tests/resources/test_master_4_2017.xlsx", 4, 2017)
# test_master_three = project_data_from_master("/home/will/code/python/analysis_engine/tests/resources/test_master_4_2018.xlsx", 4, 2018)
# test_master_four = project_data_from_master("/home/will/code/python/analysis_engine/tests/resources/test_master_4_2019.xlsx", 4, 2019)
# test_master_data = [test_master_one, test_master_two, test_master_three, test_master_fou


@pytest.fixture
def abbreviations():
    return {'Sea of Tranquility': 'SoT',
            'Apollo 11': 'A11',
            'Apollo 13': 'A13',
            'Falcon 9': 'F9',
            'Columbia': 'Columbia',
            'Mars': 'Mars'}


@pytest.fixture(scope="module")
def mst():
    test_master_data = [project_data_from_master("/home/will/code/python/"
                                                 "analysis_engine/tests/resources/test_master_4_2016.xlsx", 4, 2016),
                        project_data_from_master("/home/will/code/python/"
                                                 "analysis_engine/tests/resources/test_master_4_2017.xlsx", 4, 2017),
                        project_data_from_master("/home/will/code/python/"
                                                 "analysis_engine/tests/resources/test_master_4_2018.xlsx", 4, 2018),
                        project_data_from_master("/home/will/code/python/"
                                                 "analysis_engine/tests/resources/test_master_4_2019.xlsx", 4, 2019)]
    return Masters(test_master_data, test_master_data[0].projects)


def test_Masters_get_baseline_data(mst):
    mst.baseline_data('Re-baseline IPDC milestones')
    assert isinstance(mst.bl_index, (dict,))


def test_MilestoneData_project_dict_returns_dict(mst, abbreviations):
    mst.baseline_data("Re-baseline IPDC milestones")
    m = MilestoneData(mst, abbreviations)
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
