"""
Tests for analysis_engine
"""
import os
import datetime

from analysis_engine.data import (
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
    MilestoneData,
    milestone_chart, save_graph,
    DCA_KEYS, dca_changes_into_word, dca_changes_into_excel, DcaData, RiskData, risks_into_excel, VfMData,
    vfm_into_excel, sort_projects_by_dca, project_report_meta_data, print_out_project_milestones, put_milestones_into_wb
)

# test masters project names
sot = "Sea of Tranquility"
a11 = "Apollo 11"
a13 = "Apollo 13"
f9 = "Falcon 9"
columbia = "Columbia"
mars = "Mars"
group = [sot, a13, f9, columbia, mars]


def test_creation_of_masters_class(basic_masters_dicts, project_info):
    master = Master(basic_masters_dicts, project_info)
    assert isinstance(master.master_data, (list,))


def test_getting_baseline_data_from_masters(basic_masters_dicts, project_info):
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


# def test_check_projects_in_project_info(basic_masters_dicts, project_info_incorrect):
#     Master(basic_masters_dicts, project_info_incorrect)
#     # assert error message


def test_get_project_abbreviations(basic_masters_dicts, project_info):
    master = Master(basic_masters_dicts, project_info)
    assert master.abbreviations == {
        "Apollo 11": "A11",
        "Apollo 13": "A13",
        "Columbia": "Columbia",
        "Falcon 9": "F9",
        "Mars": "Mars",
        "Sea of Tranquility": "SoT",
    }


# assert expected error message
def test_checking_baseline_data(basic_master_wrong_baselines, project_info):
    master = Master(basic_master_wrong_baselines, project_info)
    master.check_baselines()


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


def test_project_report_meta_data(word_doc, project_info, two_masters):
    master = Master(two_masters, project_info)
    costs = CostData(master, "Falcon 9")
    m = MilestoneData(master, "Falcon 9")
    b = BenefitsData(master, "Falcon 9")
    project_report_meta_data(word_doc, costs, m, b, "Falcon 9")
    word_doc.save("resources/summary_temp_altered.docx")


def test_get_project_cost_profile(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    # master.check_baselines()
    costs = CostData(master, f9)
    assert len(costs.current_profile) == 24


def test_project_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, f9)
    cost_profile_graph(costs)


def test_project_cost_profile_chart_into_word_doc_one(
        word_doc, costs_masters, project_info
):
    master = Master(costs_masters, project_info)
    costs = CostData(master, f9)
    graph = cost_profile_graph(costs, show="No")
    change_word_doc_landscape(word_doc)
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_total_cost_profile_chart_into_word_doc_one(
        word_doc, costs_masters, project_info
):
    master = Master(costs_masters, project_info)
    costs = CostData(master, f9)
    benefits = BenefitsData(master, f9)
    graph = total_costs_benefits_bar_chart(costs, benefits, show="No")
    change_word_doc_landscape(word_doc)
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_changing_word_doc_to_landscape(word_doc):
    change_word_doc_landscape(word_doc)
    word_doc.save("resources/summary_changed_to_landscape.docx")


def test_project_cost_profile_chart_into_word_doc_many(
        word_doc, costs_masters, project_info
):
    master = Master(costs_masters, project_info)
    for p in master.current_projects:
        costs = CostData(master, p)
        graph = cost_profile_graph(costs, show="No")
        put_matplotlib_fig_into_word(word_doc, graph)
        word_doc.save("resources/summary_temp_altered.docx")


def test_get_group_cost_profile(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, master.current_projects)
    assert costs.current_profile == [
        15.45,
        932.8199999999999,
        798.1,
        406.81,
        362.8,
        344.97,
        943.07,
        1235.95,
        1362.52,
        1572.957082855212,
        1124.88,
        534.5699999999999,
        264.61,
        221.47,
        223.66,
        226.79,
        229.96,
        233.23999999999998,
        217.29999999999998,
        145.93999999999997,
        51.87,
        0.6799999999999999,
        0.6799999999999999,
        0.6799999999999999
    ]


def test_get_group_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, master.current_projects)
    cost_profile_graph(costs, title="Group Test", show="No")


def test_get_project_total_cost_calculations_for_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, f9)
    assert costs.spent == [471, 188, 188]
    assert costs.profiled == [6281, 6204, 6204]
    assert costs.unprofiled == [0, 0, 0]


def test_get_project_total_costs_benefits_bar_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, mars)
    benefits = BenefitsData(master, mars)
    total_costs_benefits_bar_chart(costs, benefits, show="No")


def test_get_group_total_cost_calculations(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, master.current_projects)
    assert costs.spent == [2929, 2210, 2210]


def test_get_group_total_cost_and_bens_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, master.current_projects)
    benefits = BenefitsData(master, master.current_projects)
    total_costs_benefits_bar_chart(costs, benefits, title="Total Group", show="No")


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
    ben = BenefitsData(master, master.current_projects)
    assert ben.delivered == [0, 0, 0]
    assert ben.profiled == [43659, 20608, 64227]
    assert ben.unprofiled == [10164, 19228, 19228]


def test_getting_benefits_profile_for_a_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    ben = BenefitsData(master, f9)
    assert ben.profiled == [-200, 240, 240]


def test_compare_changes_between_masters(basic_masters_file_paths, project_info):
    gmpp_list = get_gmpp_projects(project_info)
    wb = compare_masters(basic_masters_file_paths, gmpp_list)
    wb.save(os.path.join(os.getcwd(), "resources/cut_down_master_compared.xlsx"))


def test_get_gmpp_projects(project_info):
    gmpp_list = get_gmpp_projects(project_info)
    assert gmpp_list == ["Sea of Tranquility"]


# this method has probably now been superceded by save_graph
def test_saving_cost_profile_graph_files(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, sot)
    standard_profile(costs)
    costs = CostData(master, group)
    standard_profile(costs, title="Python", fig_size=FIGURE_STYLE[1])


# this method has probably now been superceded by save_graph
def test_saving_total_cost_benefit_graph_files(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, f9)
    benefits = BenefitsData(master, f9)
    totals_chart(costs, benefits)
    costs = CostData(master, group)
    benefits = BenefitsData(master, group)
    totals_chart(costs, benefits, title="Test Group")


def test_get_milestone_data(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, [sot, a11, a13])
    assert isinstance(milestones.current, (dict,))


def test_get_milestone_chart_data(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, [sot, a11, a13])
    assert len(milestones.key_names) == 8
    assert len(milestones.md_current) == 8
    assert len(milestones.md_last) == 8


def test_compile_milestone_chart(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, [sot, a11, a13])
    milestone_chart(milestones, title="Group Test", fig_size=FIGURE_STYLE[1], blue_line="Today")


def test_compile_milestone_chart_with_filter(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, [sot, a11, a13])
    milestones.filter_chart_info(start_date="1/1/2013", end_date="1/1/2014")
    milestone_chart(milestones, title="Group Test", fig_size=FIGURE_STYLE[1])


def test_removing_project_name_from_milestone_keys(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, sot)
    assert milestones.key_names == ["Start of Project", 'Standard A', 'Inverted Cosmonauts']


def test_putting_milestones_into_wb(milestone_masters, project_info):
    mst = Master(milestone_masters, project_info)
    milestone_data = MilestoneData(mst, a13)
    milestone_data.filter_chart_info(milestone_type=["Approval", "Delivery"])
    wb = put_milestones_into_wb(milestone_data)
    wb.save("resources/milestone_data_output_test.xlsx")


def test_saving_graph_to_word_doc_one(word_doc, milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, [sot, a11, a13])
    change_word_doc_landscape(word_doc)
    # milestones.filter_chart_info(start_date="1/1/2013", end_date="1/1/2014")
    graph = milestone_chart(milestones, title="Group Test", blue_line="Today", show="No")
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_saving_graph_to_word_doc_other(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, [sot, a11, a13])
    milestones.filter_chart_info(start_date="1/1/2013", end_date="1/1/2014")
    f = milestone_chart(milestones, title="Group Test", fig_size=FIGURE_STYLE[1], blue_line="Today")
    save_graph(f, "testing", orientation="landscape")


def test_dca_analysis(project_info, dca_masters, word_doc):
    m = Master(dca_masters, project_info)
    dca = DcaData(m)
    # assert dca.dca_count == {}
    # dca.get_changes("Q4 19/20", "Q4 18/19")
    # dca_changes_into_word(dca, word_doc)
    # word_doc.save("resources/dca_checks.docx")
    # quarter_list = ["Q4 19/20", "Q4 18/19"]
    # wb = dca_changes_into_excel(dca, quarter_list)
    # wb.save("resources/dca_print.xlsx")


def test_risk_analysis(project_info, risk_masters):
    m = Master(risk_masters, project_info)
    risk = RiskData(m, group=["Rail"])
    wb = risks_into_excel(risk)
    wb.save("resources/risks.xlsx")


def test_vfm_analysis(project_info, vfm_masters):
    m = Master(vfm_masters, project_info)
    quarter_list = ["Q1 20/21", "Q4 19/20"]
    vfm = VfMData(m, quarters=quarter_list, stage=["FBC", "OBC"])
    wb = vfm_into_excel(vfm)
    wb.save("resources/vfm.xlsx")


def test_getting_project_groups(project_info, basic_masters_dicts):
    m = Master(basic_masters_dicts, project_info)
    # assert m.dft_groups == {}
    # assert m.project_stage == {}
    assert isinstance(m.project_stage, (dict,))
    assert isinstance(m.dft_groups, (dict,))


def test_sorting_project_by_dca(project_info, dca_masters):
    rag_list = sort_projects_by_dca(dca_masters[0], group)
    assert rag_list == [('Falcon 9', 'Amber'),
                        ('Mars', 'Amber'),
                        ('Apollo 13', 'Amber/Green'),
                        ('Sea of Tranquility', 'Green'),
                        ('Columbia', 'Green')]


def test_calculating_wlc_changes(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, master.current_projects)
    assert costs.wlc_change == {'Apollo 13': {'baseline one': 0, 'last quarter': 0},
                                'Columbia': {'baseline one': -43, 'last quarter': -43},
                                'Falcon 9': {'baseline one': 5, 'last quarter': 5},
                                'Mars': {'baseline one': 0},
                                'Sea of Tranquility': {'baseline one': 54, 'last quarter': 54}}


def test_calculating_schedule_changes(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, [sot, a11, a13])
    milestones.calculate_schedule_changes()
    assert isinstance(milestones.schedule_change, (dict,))


def test_printout_of_milestones(word_doc, milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, sot)
    change_word_doc_landscape(word_doc)
    print_out_project_milestones(word_doc, milestones, sot)
    # milestones.filter_chart_info(start_date="1/1/2013", end_date="1/1/2014")
    # graph = milestone_chart(milestones, title="Group Test", blue_line="Today", show="No")
    # put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


# to write. structure for cli below.
# def test_cost_schedule_matrix():
#     m = Master(get_master_data(), get_project_information())
#     costs = CostData(m, m.current_projects)
#     milestones = MilestoneData(m, m.current_projects)
#     milestones.calculate_schedule_changes()
#     wb = cost_v_schedule_chart(milestones, costs)
#     wb.save("resources/costs_schedule_matrix.xlsx")