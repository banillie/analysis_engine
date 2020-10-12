"""
Tests for analysis_engine
"""

from data_mgmt.data import MilestoneData, Masters, current_projects
import datetime

from data_mgmt.oldegg_functions import spent_calculation
from project_analysis.p_reports import wd_heading, key_contacts, dca_table, dca_narratives


start_date = datetime.date(2020, 6, 1)
end_date = datetime.date(2022, 6, 30)


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


def test_creation_of_Masters_class(basic_master, project_info):
    projects = list(project_info.projects)
    master = Masters(basic_master, projects)
    assert isinstance(master.master_data, (list,))
    assert master.project_names == ['Mars', 'Sea of Tranquility', 'Apollo 11', 'Apollo 13', 'Falcon 9', 'Columbia']


def test_getting_baseline_data_from_Masters(basic_master, project_info):
    projects = list(project_info.projects)
    master = Masters(basic_master, projects)
    assert isinstance(master.bl_index, (dict,))
    assert master.bl_index["ipdc_milestones"]["Sea of Tranquility"] == [0, 1]
    assert master.bl_index["ipdc_costs"]["Apollo 11"] == [0, 1, 2]


def test_calculating_spent(spent_master):
    spent = spent_calculation(spent_master, "Sea of Tranquility")
    assert spent == 439.9

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
