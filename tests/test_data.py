"""
Tests for analysis_engine
"""

from data_mgmt.data import Master, CostData, spent_calculation, wd_heading, \
    key_contacts, dca_table, dca_narratives, project_cost_profile_graph, year_cost_profile_chart, \
    group_cost_profile_graph, total_costs_benefits_bar_chart


def test_creation_of_Masters_class(basic_master, project_info):
    master = Master(basic_master, project_info)
    assert isinstance(master.master_data, (list,))


def test_getting_baseline_data_from_Masters(basic_master, project_info):
    master = Master(basic_master, project_info)
    assert isinstance(master.bl_index, (dict,))
    assert master.bl_index["ipdc_milestones"]["Sea of Tranquility"] == [0, 1]
    assert master.bl_index["ipdc_costs"]["Apollo 11"] == [0, 1, 2]
    assert master.bl_index["ipdc_costs"]["Columbia"] == [0, 1, 2]


def test_get_current_project_names(basic_master, project_info):
    master = Master(basic_master, project_info)
    assert master.current_projects == ['Sea of Tranquility', 'Apollo 11', 'Apollo 13', 'Falcon 9', 'Columbia']


def test_check_projects_in_project_info(basic_master, project_info_incorrect):
    Master(basic_master, project_info_incorrect)
    # assert error message


def test_checking_baseline_data(basic_master_wrong_baselines, project_info):
    master = Master(basic_master_wrong_baselines, project_info)
    master.check_baselines()
    # assert expected error message


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


def test_word_doc_contacts(word_doc, project_info, contact_master):
    master = Master(contact_master, project_info)
    key_contacts(word_doc, master, 'Apollo 13')
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_dca_table(word_doc, project_info, dca_masters):
    master = Master(dca_masters, project_info)
    dca_table(word_doc, master, 'Falcon 9')
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_dca_narratives(word_doc, project_info, dca_masters):
    master = Master(dca_masters, project_info)
    dca_narratives(word_doc, master, 'Falcon 9')
    word_doc.save("resources/summary_temp_altered.docx")


def test_get_project_cost_profile(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_profile_project('Falcon 9', 'ipdc_costs')
    assert costs.current_profile_project == [0, 0, 177.49, 245, 411.3, 443.2, 728.1, 1046.6, 1441, 1315, 395.84, 0]
    assert costs.last_profile_project == [0, 78.4, 165, 216.1, 323.95, 825.71, 909.19, 1216.59, 1141.08, 706.25, 0, 0]
    assert costs.baseline_profile_one_project == [0, 78.4, 165, 216.1, 323.95, 825.71, 909.19, 1216.59, 1141.08, 706.25, 0, 0]
    assert costs.rdel_profile_project == [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    assert costs.cdel_profile_project == [0, 0, 177.49, 245, 411.3, 443.2, 728.1, 1046.6, 1441, 1315, 395.84, 0]
    assert costs.ngov_profile_project == [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]


def test_project_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    costs.get_profile_project('Falcon 9', 'ipdc_costs')
    project_cost_profile_graph(costs)


def test_project_cost_profile_chart_into_word_doc_one(word_doc, costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    costs.get_profile_project('Falcon 9', 'ipdc_costs')
    year_cost_profile_chart(word_doc, costs)
    word_doc.save("resources/summary_temp_altered.docx")


def test_project_cost_profile_chart_into_word_doc_many(word_doc, costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    for p in master.current_projects:
        costs.get_profile_project(p, 'ipdc_costs')
        year_cost_profile_chart(word_doc, costs)
        word_doc.save("resources/summary_temp_altered.docx")


def test_get_group_cost_profile(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_profile_all('ipdc_costs')
    assert costs.current_profile == [0, 0, 265, 266, 412, 444, 729, 1047, 1442, 1316, 396, 1]


def test_get_group_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_profile_all('ipdc_costs')
    group_cost_profile_graph(costs, 'Group Test')


def test_get_total_cost_calculations_for_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_cost_totals_project('Falcon 9', 'ipdc_costs')
    assert costs.spent == [188, 110, 110]
    assert costs.profiled == [6204, 5582, 5582]
    assert costs.unprofiled == [0, 0, 0]

def test_get_total_costs_benefits_bar_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_cost_totals_project('Apollo 13', 'ipdc_costs')
    total_costs_benefits_bar_chart(costs)


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
