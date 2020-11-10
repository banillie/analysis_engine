"""
Tests for analysis_engine
"""
from data_mgmt.data import Master, CostData, spent_calculation, wd_heading, \
    key_contacts, dca_table, dca_narratives, project_cost_profile_graph, year_cost_profile_chart, \
    group_cost_profile_graph, total_costs_benefits_bar_chart_project, total_costs_benefits_bar_chart_group, \
    run_get_old_fy_data, run_place_old_fy_data_into_masters, put_key_change_master_into_dict, run_change_keys, \
    BenefitsData


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
 #    assert costs.current_profile_project == [0,
 # 0,
 # 177.49,
 # 245,
 # 411.3,
 # 443.2,
 # 728.1,
 # 1046.6,
 # 1441,
 # 1315,
 # 395.84,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0]
 #    assert costs.last_profile_project == [0,
 # 78.4,
 # 165,
 # 216.1,
 # 323.95,
 # 825.71,
 # 909.19,
 # 1216.59,
 # 1141.08,
 # 706.25,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0,
 # 0]
    # assert costs.baseline_profile_one_project == [0, 78.4, 165, 216.1, 323.95, 825.71, 909.19, 1216.59, 1141.08, 706.25, 0, 0]
    # assert costs.rdel_profile_project == [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    # assert costs.cdel_profile_project == [0, 0, 177.49, 245, 411.3, 443.2, 728.1, 1046.6, 1441, 1315, 395.84, 0]
    # assert costs.ngov_profile_project == [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]


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
    costs.get_profile_group('ipdc_costs')
    # assert costs.current_profile == [0, 0, 265, 266, 412, 444, 729, 1047, 2441, 3316, 396, 1]


def test_get_group_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_profile_group('ipdc_costs')
    group_cost_profile_graph(costs, 'Group Test')


def test_get_project_total_cost_calculations_for_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_cost_totals_project('Falcon 9', 'ipdc_costs')
    # assert costs.spent == [471, 188, 188]
    # assert costs.profiled == [6204, 5582, 5582]
    # assert costs.unprofiled == [0, 0, 0]


def test_get_project_total_costs_benefits_bar_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    benefits = BenefitsData(master)
    costs.get_cost_totals_project('Mars', 'ipdc_costs')
    benefits.get_ben_totals_project('Mars', 'ipdc_costs')
    total_costs_benefits_bar_chart_project(costs, benefits)


def test_get_group_total_cost_calculations(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_cost_totals_group('ipdc_costs')
    assert costs.spent == [468, 2210, 2210]


def test_get_group_total_cost_and_bens_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    bens = BenefitsData(master)
    costs.get_cost_totals_group('ipdc_costs')
    bens.get_ben_totals_group('ipdc_benefits')
    total_costs_benefits_bar_chart_group(costs, bens, 'Total Group')


def test_put_change_keys_into_a_dict(change_log):
    keys_dict = put_key_change_master_into_dict(change_log)
    assert isinstance(keys_dict, (dict,))


def test_altering_master_wb_file_key_names(change_log, list_cost_masters_files, list_test_masters_files):
    keys_dict = put_key_change_master_into_dict(change_log)
    run_change_keys(list_cost_masters_files, keys_dict)
    run_change_keys(list_test_masters_files, keys_dict)


def test_get_old_fy_cost_data(list_test_masters_files, project_group_id_path):
    run_get_old_fy_data(list_test_masters_files, project_group_id_path)


def test_placing_old_fy_cost_data_into_master_wbs(list_cost_masters_files, project_group_id_path):
    run_place_old_fy_data_into_masters(list_cost_masters_files, project_group_id_path)


def test_getting_benefits_profile_for_a_group(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    ben = BenefitsData(master)
    ben.get_ben_totals_group('ipdc_benefits')
    assert ben.delivered == [0, 0, 0]
    assert ben.profiled == [-43659, -20608, -64227]
    assert ben.unprofiled == [53823, 39836, 83455]


def test_getting_benefits_profile_for_a_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    ben = BenefitsData(master)
    ben.get_ben_totals_project('Falcon 9', 'ipdc_benefits')
    assert ben.profiled_project == [-200, 240, 240]
