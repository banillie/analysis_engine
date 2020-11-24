"""
Tests for analysis_engine
"""
import os

from data_mgmt.data import (
    Master,
    CostData,
    spent_calculation,
    wd_heading,
    key_contacts,
    dca_table,
    dca_narratives,
    put_matplotlib_fig_into_word,
    cost_profile_graph,
    total_costs_benefits_bar_chart,
    run_get_old_fy_data,
    run_place_old_fy_data_into_masters,
    put_key_change_master_into_dict,
    run_change_keys,
    BenefitsData,
    compare_masters,
    get_gmpp_projects,
    standard_profile,
    totals_chart,
    change_word_doc_landscape,
    FIGURE_STYLE,
    MilestoneData, milestone_chart,
)

# test masters project names
sot = "Sea of Tranquility"
a11 = "Apollo 11"
a13 = "Apollo 13"
f9 = "Falcon 9"
columbia = "Columbia"
mars = "Mars"
group = [sot, a13, f9, columbia, mars]


def test_creation_of_Masters_class(basic_masters_dicts, project_info):
    master = Master(basic_masters_dicts, project_info)
    assert isinstance(master.master_data, (list,))


def test_getting_baseline_data_from_Masters(basic_masters_dicts, project_info):
    master = Master(basic_masters_dicts, project_info)
    assert isinstance(master.bl_index, (dict,))
    assert master.bl_index["ipdc_milestones"]["Sea of Tranquility"] == [0, 1]
    assert master.bl_index["ipdc_costs"]["Apollo 11"] == [0, 1, 2]
    assert master.bl_index["ipdc_costs"]["Columbia"] == [0, 1, 2]


def test_get_current_project_names(basic_masters_dicts, project_info):
    master = Master(basic_masters_dicts, project_info)
    assert master.current_projects == [
        "Sea of Tranquility",
        "Apollo 11",
        "Apollo 13",
        "Falcon 9",
        "Columbia",
    ]


def test_check_projects_in_project_info(basic_masters_dicts, project_info_incorrect):
    Master(basic_masters_dicts, project_info_incorrect)
    # assert error message


def test_checking_baseline_data(basic_master_wrong_baselines, project_info):
    master = Master(basic_master_wrong_baselines, project_info)
    master.check_baselines()
    # assert expected error message


def test_calculating_spent(spent_master):
    spent = spent_calculation(spent_master, "Sea of Tranquility")
    assert spent == 439.9


def test_open_word_doc(word_doc):
    word_doc.add_paragraph(
        "Because i'm still in love with you I want to see you dance again, "
        "because i'm still in love with you on this harvest moon"
    )
    word_doc.save("resources/summary_temp_altered.docx")
    var = word_doc.paragraphs[1].text
    assert (
        "Because i'm still in love with you I want to see you dance again, "
        "because i'm still in love with you on this harvest moon" == var
    )


def test_word_doc_heading(word_doc, project_info):
    wd_heading(word_doc, project_info, "Apollo 11")
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_contacts(word_doc, project_info, contact_master):
    master = Master(contact_master, project_info)
    key_contacts(word_doc, master, "Apollo 13")
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_dca_table(word_doc, project_info, dca_masters):
    master = Master(dca_masters, project_info)
    dca_table(word_doc, master, "Falcon 9")
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_dca_narratives(word_doc, project_info, dca_masters):
    master = Master(dca_masters, project_info)
    dca_narratives(word_doc, master, "Falcon 9")
    word_doc.save("resources/summary_temp_altered.docx")


def test_get_project_cost_profile(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    master.check_baselines()
    costs = CostData(master)
    costs.get_cost_profile("Falcon 9", "ipdc_costs")
    assert len(costs.current_profile) == 24


def test_project_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    costs.get_cost_profile("Falcon 9", "ipdc_costs")
    cost_profile_graph(FIGURE_STYLE[2], costs, "Falcon 9")


def test_project_cost_profile_chart_into_word_doc_one(
    word_doc, costs_masters, project_info
):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    costs.get_cost_profile("Falcon 9", "ipdc_costs")
    fig_style = "half horizontal"
    graph = cost_profile_graph(fig_style, costs)
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_total_cost_profile_chart_into_word_doc_one(
    word_doc, costs_masters, project_info
):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    benefits = BenefitsData(master)
    costs.get_cost_totals("Falcon 9", "ipdc_costs")
    benefits.get_ben_totals("Falcon 9", "ipdc_benefits")
    fig_size = FIGURE_STYLE[2]
    graph = total_costs_benefits_bar_chart(fig_size, costs, benefits)
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_changing_word_doc_to_landscape(word_doc):
    change_word_doc_landscape(word_doc)
    word_doc.save("resources/summary_changed_to_landscape.docx")


def test_project_cost_profile_chart_into_word_doc_many(
    word_doc, costs_masters, project_info
):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    for p in master.current_projects:
        costs.get_cost_profile(p, "ipdc_costs")
        graph = cost_profile_graph(FIGURE_STYLE[2], costs)
        put_matplotlib_fig_into_word(word_doc, graph)
        word_doc.save("resources/summary_temp_altered.docx")


def test_get_group_cost_profile(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    costs.get_cost_profile(master.current_projects, "ipdc_costs")
    assert costs.current_profile == [
        0,
        933,
        798,
        407,
        363,
        345,
        943,
        1236,
        1363,
        1573,
        1125,
        535,
        265,
        221,
        224,
        227,
        230,
        233,
        217,
        146,
        52,
        1,
        1,
        1,
    ]


def test_get_group_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    costs.get_cost_profile(master.current_projects, "ipdc_costs")
    cost_profile_graph(FIGURE_STYLE[2], costs, "Group Test")


def test_get_project_total_cost_calculations_for_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    costs.get_cost_totals("Falcon 9", "ipdc_costs")
    assert costs.spent == [471, 188, 188]
    assert costs.profiled == [6281, 6204, 6204]
    assert costs.unprofiled == [0, 0, 0]


def test_get_project_total_costs_benefits_bar_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    benefits = BenefitsData(master)
    costs.get_cost_totals("Mars", "ipdc_costs")
    benefits.get_ben_totals("Mars", "ipdc_benefits")
    total_costs_benefits_bar_chart(FIGURE_STYLE[2], costs, benefits)


def test_get_group_total_cost_calculations(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    costs.get_cost_totals(master.current_projects, "ipdc_costs")
    assert costs.spent == [2929, 2210, 2210]


def test_get_group_total_cost_and_bens_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    bens = BenefitsData(master)
    costs.get_cost_totals(master.current_projects, "ipdc_costs")
    bens.get_ben_totals(master.current_projects, "ipdc_benefits")
    total_costs_benefits_bar_chart(FIGURE_STYLE[2], costs, bens, "Total Group")


def test_put_change_keys_into_a_dict(change_log):
    keys_dict = put_key_change_master_into_dict(change_log)
    assert isinstance(keys_dict, (dict,))


def test_altering_master_wb_file_key_names(change_log, list_cost_masters_files):
    keys_dict = put_key_change_master_into_dict(change_log)
    run_change_keys(list_cost_masters_files, keys_dict)


def test_get_old_fy_cost_data(list_cost_masters_files, project_group_id_path):
    run_get_old_fy_data(list_cost_masters_files, project_group_id_path)


def test_placing_old_fy_cost_data_into_master_wbs(
    list_cost_masters_files, project_old_fy_path
):
    run_place_old_fy_data_into_masters(list_cost_masters_files, project_old_fy_path)


def test_getting_benefits_profile_for_a_group(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    ben = BenefitsData(master)
    ben.get_ben_totals(master.current_projects, "ipdc_benefits")
    assert ben.delivered == [0, 0, 0]
    assert ben.profiled == [43659, 20608, 64227]
    assert ben.unprofiled == [10164, 19228, 19228]


def test_getting_benefits_profile_for_a_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    ben = BenefitsData(master)
    ben.get_ben_totals("Falcon 9", "ipdc_benefits")
    assert ben.profiled == [-200, 240, 240]


def test_compare_changes_between_masters(basic_masters_file_paths, project_info):
    gmpp_list = get_gmpp_projects(project_info)
    wb = compare_masters(basic_masters_file_paths, gmpp_list)
    wb.save(os.path.join(os.getcwd(), "resources/cut_down_master_compared.xlsx"))


def test_get_gmpp_projects(project_info):
    gmpp_list = get_gmpp_projects(project_info)
    assert gmpp_list == ["Sea of Tranquility"]


def test_saving_cost_profile_graph_files(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    fig_style = "half horizontal"
    standard_profile(fig_style, costs, sot)
    standard_profile(fig_style, costs, group, "Python")


def test_saving_total_cost_benefit_graph_files(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master)
    benefits = BenefitsData(master)
    fig_style = FIGURE_STYLE[2]
    totals_chart(fig_style, costs, benefits, f9)
    totals_chart(fig_style, costs, benefits, group, "Matplotlib")


def test_get_milestone_data(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master)
    milestones.get_milestones([sot, a11, a13], "ipdc_milestones")
    assert isinstance(milestones.current, (dict,))


def test_get_milestone_chart_data(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master)
    milestones.get_milestones([sot, a11, a13], "ipdc_milestones")
    milestones.get_chart_info()
    assert len(milestones.key_names) == 5
    assert len(milestones.md_current) == 5
    assert len(milestones.md_last) == 5


def test_compile_milestone_chart(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master)
    milestones.get_milestones([sot, a11, a13], "ipdc_milestones")
    milestones.get_chart_info()
    milestone_chart("half horizontal", milestones, "testing")