"""
Tests for analysis_engine
"""

from data_mgmt.data import MilestoneData, Masters, current_projects, CostData, project_cost_profile_graph, \
    group_cost_profile_graph
import datetime

from data_mgmt.oldegg_functions import spent_calculation
from project_analysis.p_reports import wd_heading, key_contacts, dca_table, dca_narratives, year_cost_profile_chart


def test_creation_of_Masters_class(basic_master, project_info):
    projects = list(project_info.projects)
    master = Masters(basic_master, projects)
    assert isinstance(master.master_data, (list,))
    assert master.project_names == ['Mars', 'Sea of Tranquility', 'Apollo 11', 'Apollo 13', 'Falcon 9', 'Columbia']


def test_getting_baseline_data_from_Masters(basic_master, project_info):
    projects = list(project_info.projects)
    master = Masters(basic_master, projects)
    master.baseline_data()
    assert isinstance(master.bl_index, (dict,))
    assert master.bl_index["ipdc_milestones"]["Sea of Tranquility"] == [0, 1]
    assert master.bl_index["ipdc_costs"]["Apollo 11"] == [0, 1, 2]
    assert master.bl_index["ipdc_costs"]["Columbia"] == [0, 1, 2]


def test_calculating_spent(spent_master):
    spent = spent_calculation(spent_master, "Sea of Tranquility")
    assert spent == 439.9


def test_open_word_doc(word_doc):
    word_doc.add_paragraph("Because i'm still in love with you I want to see you dance again, "
                           "because i'm still in love with you on this harvest moon")
    word_doc.save("resources/summary_temp_altered.docx")
    var = word_doc.paragraphs[1].text
    assert "Because i'm still in love with you I want to see you dance again, " \
           "because i'm still in love with you on this harvest moon" == var


def test_word_doc_heading(word_doc, project_info):
    wd_heading(word_doc, project_info, 'Apollo 11')
    word_doc.save("resources/summary_temp_altered.docx")


def test_list_of_current_projects(project_info):
    live_projects = current_projects(project_info)
    assert live_projects == ['Sea of Tranquility', 'Apollo 11', 'Apollo 13', 'Falcon 9']


def test_word_doc_contacts(word_doc, project_info, contact_master):
    live_projects = current_projects(project_info)
    master = Masters(contact_master, live_projects)
    key_contacts(word_doc, master, 'Apollo 13')
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_dca_table(word_doc, project_info, dca_masters):
    live_projects = current_projects(project_info)
    master = Masters(dca_masters, live_projects)
    dca_table(word_doc, master, 'Falcon 9')
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_dca_narratives(word_doc, project_info, dca_masters):
    live_projects = current_projects(project_info)
    master = Masters(dca_masters, live_projects)
    dca_narratives(word_doc, master, 'Falcon 9')
    word_doc.save("resources/summary_temp_altered.docx")


def test_get_project_cost_profile(costs_masters, project_info):
    live_projects = current_projects(project_info)
    master = Masters(costs_masters, live_projects)
    costs = CostData(master)
    costs.get_profile_project('Falcon 9', 'ipdc_costs')
    assert costs.current_profile_project == [0, 0, 177.49, 245, 411.3, 443.2, 728.1, 1046.6, 1441, 1315, 395.84, 0]
    assert costs.last_profile_project == [0, 78.4, 165, 216.1, 323.95, 825.71, 909.19, 1216.59, 1141.08, 706.25, 0, 0]
    assert costs.baseline_profile_one_project == [0, 78.4, 165, 216.1, 323.95, 825.71, 909.19, 1216.59, 1141.08, 706.25, 0, 0]
    assert costs.rdel_profile_project == [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    assert costs.cdel_profile_project == [0, 0, 177.49, 245, 411.3, 443.2, 728.1, 1046.6, 1441, 1315, 395.84, 0]
    assert costs.ngov_profile_project == [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]


def test_project_cost_profile_chart(costs_masters, project_info):
    live_projects = current_projects(project_info)
    master = Masters(costs_masters, live_projects)
    costs = CostData(master)
    costs.get_profile_project('Falcon 9', 'ipdc_costs')
    project_cost_profile_graph(costs)


def test_project_cost_profile_chart_into_word_doc_one(word_doc, costs_masters):
    live_projects = ['Falcon 9', 'Columbia', 'Apollo 13']
    master = Masters(costs_masters, live_projects)
    costs = CostData(master)
    costs.get_profile_project('Falcon 9', 'ipdc_costs')
    year_cost_profile_chart(word_doc, costs)
    word_doc.save("resources/summary_temp_altered.docx")


def test_project_cost_profile_chart_into_word_doc_many(word_doc, costs_masters):
    live_projects = ['Falcon 9', 'Columbia', 'Apollo 13']
    master = Masters(costs_masters, live_projects)
    costs = CostData(master)
    for p in live_projects:
        costs.get_profile_project(p, 'ipdc_costs')
        year_cost_profile_chart(word_doc, costs)
        word_doc.save("resources/summary_temp_altered.docx")


def test_get_group_cost_profile(costs_masters, project_info):
    live_projects = ['Falcon 9', 'Columbia', 'Apollo 13']
    master = Masters(costs_masters, live_projects)
    costs = CostData(master)
    costs.get_profile_all('ipdc_costs')
    assert costs.current_profile == [0, 0, 230, 266, 412, 444, 729, 1047, 1442, 1316, 396, 1]


def test_get_group_cost_profile_chart(costs_masters, project_info):
    live_projects = ['Falcon 9', 'Columbia', 'Apollo 13']
    master = Masters(costs_masters, live_projects)
    costs = CostData(master)
    costs.get_profile_all('ipdc_costs')
    group_cost_profile_graph(costs, 'Group Test')



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
