"""
Tests for analysis_engine
"""
import csv
import os
import datetime
import pickle
from matplotlib import pyplot as plt

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
    milestone_chart,
    save_graph,
    DCA_KEYS,
    dca_changes_into_word,
    dca_changes_into_excel,
    DcaData,
    RiskData,
    risks_into_excel,
    VfMData,
    vfm_into_excel,
    sort_projects_by_dca,
    project_report_meta_data,
    print_out_project_milestones,
    put_milestones_into_wb,
    Pickle,
    open_pickle_file,
    financial_dashboard,
    schedule_dashboard,
    benefits_dashboard,
    overall_dashboard,
    DandelionData,
    dandelion_data_into_wb,
    # run_dandelion_matplotlib_chart,
    cost_v_schedule_chart_into_wb,
    cost_profile_into_wb,
    data_query_into_wb,
    get_data_query_key_names,
    remove_project_name_from_milestone_key, get_sp_data, cost_stackplot_graph, get_group,
    make_a_dandelion_manual, make_a_dandelion_auto,
)

# test masters project names

sot = "Sea of Tranquility"
a11 = "Apollo 11"
a13 = "Apollo 13"
f9 = "Falcon 9"
columbia = "Columbia"
mars = "Mars"
group = [sot, a13, f9, columbia, mars]


def test_master_in_a_pickle(basic_masters_dicts, project_info):
    master = Master(basic_masters_dicts, project_info)
    path_str = str("{0}/resources/test_master".format(os.path.join(os.getcwd())))
    mickle = Pickle(master, path_str)
    assert str(mickle.master.master_data[0].quarter) == "Q4 16/17"


def test_opening_a_pickle(basic_pickle):
    mickle = open_pickle_file(basic_pickle)
    assert str(mickle.master_data[0].quarter) == "Q4 16/17"


def test_creation_of_masters_class(basic_masters_dicts, project_info):
    master = Master(basic_masters_dicts, project_info)
    assert isinstance(master.master_data, (list,))


def test_getting_baseline_data_from_masters(basic_masters_dicts, project_info):
    master = Master(basic_masters_dicts, project_info)
    assert isinstance(master.bl_index, (dict,))
    assert master.bl_index["ipdc_milestones"]["Sea of Tranquility"] == [0, 1]
    assert master.bl_index["ipdc_costs"]["Apollo 11"] == [0, 1, 0, 2]
    assert master.bl_index["ipdc_costs"]["Columbia"] == [0, 1, 0, 2]


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
        "Apollo 11": {"abb": "A11", "full name": "Apollo 11"},
        "Apollo 13": {"abb": "A13", "full name": "Apollo 13"},
        "Columbia": {"abb": "Columbia", "full name": "Columbia"},
        "Falcon 9": {"abb": "F9", "full name": "Falcon 9"},
        "Mars": {"abb": "Mars", "full name": "Mars"},
        "Sea of Tranquility": {"abb": "SoT", "full name": "Sea of Tranquility"},
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
    c = CostData(master, quarter=["standard"])
    m = MilestoneData(master, quarter=["standard"])
    b = BenefitsData(master, quarter=["standard"])
    project_report_meta_data(word_doc, c, m, b, "Falcon 9")
    word_doc.save("resources/summary_temp_altered.docx")


def test_get_project_cost_profile(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, group=[mars], baseline=['standard'])
    assert len(costs.c_profiles['current']['prof']) == 24


def test_project_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, remove=[f9], group=group, baseline=['standard'])
    cost_profile_graph(costs, chart=False)


def test_project_cost_profile_into_wb(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, baseline=['standard'])
    wb = cost_profile_into_wb(costs)
    wb.save("resources/test_cost_profile_output.xlsx")


def test_project_cost_profile_chart_into_word_doc_one(
    word_doc, costs_masters, project_info
):
    master = Master(costs_masters, project_info)
    costs = CostData(master, group=[f9], baseline=["standard"])
    graph = cost_profile_graph(costs, show=False)
    change_word_doc_landscape(word_doc)
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_get_project_total_costs_benefits_bar_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, baseline=["standard"])
    benefits = BenefitsData(master, baseline=["standard"])
    total_costs_benefits_bar_chart(costs, benefits, chart=True)


def test_total_cost_profile_chart_into_word_doc_one(
    word_doc, costs_masters, project_info
):
    master = Master(costs_masters, project_info)
    costs = CostData(master, group=[f9], baseline=["standard"])
    benefits = BenefitsData(master, group=[f9], baseline=["standard"])
    graph = total_costs_benefits_bar_chart(costs, benefits, show="No")
    change_word_doc_landscape(word_doc)
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_changing_word_doc_to_landscape(word_doc):
    change_word_doc_landscape(word_doc)
    word_doc.save("resources/summary_changed_to_landscape.docx")


def test_get_group_cost_profile(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, group=[master.current_projects], quarter=["standard"])
    assert costs.c_profiles["Q1 20/21"]["prof"] == [
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
        0.6799999999999999,
    ]


def test_get_group_cost_profile_chart(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, quarter=["standard"])
    cost_profile_graph(costs, chart=True)


def test_get_stackplot_costs_chart(costs_masters, project_info):
    m = Master(costs_masters, project_info)
    sp = get_sp_data(m, group, "Q1 20/21", type="cat")
    cost_stackplot_graph(sp)


def test_get_project_total_cost_calculations_for_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, group=[f9], baseline=["standard"])
    assert costs.c_totals["current"]["spent"] == 471
    assert costs.c_totals["current"]["prof"] == 6281
    assert costs.c_totals["current"]["unprof"] == 0


def test_get_group_total_cost_calculations(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, group=master.current_projects, quarter=["standard"])
    assert costs.c_totals["Q1 20/21"]["spent"] == 3102
    assert costs.c_totals["Q4 19/20"]["spent"] == 2210


## no longer neccesary. covered by another test.
# def test_get_group_total_cost_and_bens_chart(costs_masters, project_info):
#     master = Master(costs_masters, project_info)
#     costs = CostData(master, group=[master.current_projects])
#     benefits = BenefitsData(master, master.current_projects)
#     total_costs_benefits_bar_chart(costs, benefits, title="Total Group", show="No")


def test_put_change_keys_into_a_dict(change_log):
    keys_dict = put_key_change_master_into_dict(change_log)
    assert isinstance(keys_dict, (dict,))


def test_altering_master_wb_file_key_names(change_log, list_cost_masters_files):
    keys_dict = put_key_change_master_into_dict(change_log)
    run_change_keys(list_cost_masters_files, keys_dict)


# def test_get_old_fy_cost_data(list_cost_masters_files, project_group_id_path):
#     run_get_old_fy_data(list_cost_masters_files, project_group_id_path)
#
#
# def test_placing_old_fy_cost_data_into_master_wbs(
#     list_cost_masters_files, project_old_fy_path
# ):
#     run_place_old_fy_data_into_masters(list_cost_masters_files, project_old_fy_path)


def test_getting_benefits_profile_for_a_group(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    ben = BenefitsData(master, quarter=["standard"])
    assert ben.b_totals["Q1 20/21"]["delivered"] == 0
    assert ben.b_totals["Q1 20/21"]["prof"] == 43659
    assert ben.b_totals["Q1 20/21"]["unprof"] == 10164


def test_getting_benefits_profile_for_a_project(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    ben = BenefitsData(master, group=[f9], baseline=["all"])
    assert ben.b_totals["current"]["cat_prof"] == [0, 0, 0, -200]


def test_compare_changes_between_masters(basic_masters_file_paths, project_info):
    gmpp_list = get_gmpp_projects(project_info)
    wb = compare_masters(basic_masters_file_paths, gmpp_list)
    wb.save(os.path.join(os.getcwd(), "resources/cut_down_master_compared.xlsx"))

## nolonger using this meta data. gmpp taken from elsewhere
# def test_get_gmpp_projects(project_info):
#     gmpp_list = get_gmpp_projects(project_info)
#     assert gmpp_list == ["Sea of Tranquility"]


# this method has probably now been superceded by save_graph
# def test_saving_cost_profile_graph_files(costs_masters, project_info):
#     master = Master(costs_masters, project_info)
#     costs = CostData(master, sot)
#     standard_profile(costs)
#     costs = CostData(master, group)
#     standard_profile(costs, title="Python", fig_size=FIGURE_STYLE[1])


# this method has probably now been superceded by save_graph
# def test_saving_total_cost_benefit_graph_files(costs_masters, project_info):
#     master = Master(costs_masters, project_info)
#     costs = CostData(master, f9)
#     benefits = BenefitsData(master, f9)
#     totals_chart(costs, benefits)
#     costs = CostData(master, group)
#     benefits = BenefitsData(master, group)
#     totals_chart(costs, benefits, title="Test Group")


def test_get_milestone_data_bl(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, group=[sot, a11, a13], baseline=["all"])
    assert isinstance(milestones.milestone_dict["current"], (dict,))


def test_get_milestone_data_all(milestone_masters, project_info):
    m = Master(milestone_masters, project_info)
    milestones = MilestoneData(m, quarter=["Q4 19/20", "Q4 18/19"])
    assert isinstance(milestones.milestone_dict["Q4 19/20"], (dict,))


def test_get_milestone_chart_data(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, group=[sot, a11, a13], baseline=["standard"])
    assert (
        len(milestones.sorted_milestone_dict[milestones.iter_list[0]]["g_dates"]) == 11
    )
    assert (
        len(milestones.sorted_milestone_dict[milestones.iter_list[1]]["g_dates"]) == 11
    )
    assert (
        len(milestones.sorted_milestone_dict[milestones.iter_list[2]]["g_dates"]) == 11
    )


def test_compile_milestone_chart(milestone_masters, project_info, word_doc):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, group=[sot], quarter=["Q4 19/20", "Q4 18/19"])
    graph = milestone_chart(
        milestones, title="Group Test", fig_size=FIGURE_STYLE[1], blue_line="Today"
    )
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_compile_milestone_chart_with_filter(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, group=[sot, a11, a13], baseline=["standard"])
    milestones.filter_chart_info(dates=["1/1/2013", "1/1/2014"])
    milestone_chart(milestones, title="Group Test", fig_size=FIGURE_STYLE[1])


def test_removing_project_name_from_milestone_keys(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, group=[sot], baseline=["all"])
    key_names = milestones.sorted_milestone_dict["current"]["names"]
    key_names = remove_project_name_from_milestone_key("SoT", key_names)
    assert key_names == [
        "Start of Project",
        "Standard A",
        "Inverted Cosmonauts",
        "Start of Construction/build",
    ]


def test_putting_milestones_into_wb(milestone_masters, project_info):
    mst = Master(milestone_masters, project_info)
    milestones = MilestoneData(mst, group=[sot], baseline=["standard"])
    wb = put_milestones_into_wb(milestones)
    wb.save("resources/milestone_data_output_test.xlsx")


def test_saving_graph_to_word_doc_one(word_doc, milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, group=[sot, a11, a13], baseline=["standard"])
    change_word_doc_landscape(word_doc)
    graph = milestone_chart(milestones, title="Group Test", blue_line="Today")
    put_matplotlib_fig_into_word(word_doc, graph, size=2)
    word_doc.save("resources/summary_temp_altered.docx")


#  hashed out. not saving to test/resources
# def test_saving_graph_to_word_doc_other(milestone_masters, project_info):
#     master = Master(milestone_masters, project_info)
#     milestones = MilestoneData(master, [sot, a11, a13])
#     milestones.filter_chart_info(start_date="1/1/2013", end_date="1/1/2014")
#     f = milestone_chart(milestones, title="Group Test", fig_size=FIGURE_STYLE[1], blue_line="Today")
#     save_graph(f, "testing", orientation="landscape")


def test_dca_analysis(project_info, dca_masters, word_doc):
    m = Master(dca_masters, project_info)
    dca = DcaData(m, quarter=["standard"])
    wb = dca_changes_into_excel(dca)
    wb.save("resources/dca_print.xlsx")


def test_speedial_print_out(project_info, dca_masters, word_doc):
    m = Master(dca_masters, project_info)
    dca = DcaData(m, baseline=["standard"])
    dca.get_changes()
    dca_changes_into_word(dca, word_doc)
    word_doc.save("resources/dca_checks.docx")


def test_risk_analysis(project_info, risk_masters):
    m = Master(risk_masters, project_info)
    risk = RiskData(m, group=["Rail"], quarter=["standard"])
    wb = risks_into_excel(risk)
    wb.save("resources/risks.xlsx")


def test_vfm_analysis(project_info, vfm_masters):
    m = Master(vfm_masters, project_info)
    vfm = VfMData(m, quarter=['standard'])
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
    assert rag_list == [
        ("Falcon 9", "Amber"),
        ("Mars", "Amber"),
        ("Apollo 13", "Amber/Green"),
        ("Sea of Tranquility", "Green"),
        ("Columbia", "Green"),
    ]


def test_calculating_wlc_changes(costs_masters, project_info):
    master = Master(costs_masters, project_info)
    costs = CostData(master, group=[master.current_projects], baseline=["all"])
    costs.calculate_wlc_change()
    assert costs.wlc_change == {
        "Apollo 13": {"baseline one": 0, "last quarter": 0},
        "Columbia": {"baseline one": -43, "last quarter": -43},
        "Falcon 9": {"baseline one": 5, "last quarter": 5},
        "Mars": {"baseline one": 0},
        "Sea of Tranquility": {"baseline one": 54, "last quarter": 54},
    }


def test_calculating_schedule_changes(milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, group=[sot, a11, a13])
    milestones.get_milestones()
    milestones.get_chart_info()
    milestones.calculate_schedule_changes()
    assert isinstance(milestones.schedule_change, (dict,))


def test_printout_of_milestones(word_doc, milestone_masters, project_info):
    master = Master(milestone_masters, project_info)
    milestones = MilestoneData(master, group=[sot], baseline=["standard"])
    change_word_doc_landscape(word_doc)
    print_out_project_milestones(word_doc, milestones, sot)
    word_doc.save("resources/summary_temp_altered.docx")


def test_cost_schedule_matrix(two_masters, project_info):
    m = Master(two_masters, project_info)
    costs = CostData(m, group=m.current_projects, quarters=["standard"])
    milestones = MilestoneData(m, group=m.current_projects)
    milestones.get_milestones()
    milestones.get_chart_info()
    milestones.calculate_schedule_changes()
    wb = cost_v_schedule_chart_into_wb(milestones, costs)
    wb.save("resources/test_costs_schedule_matrix.xlsx")


def test_financial_dashboard(costs_masters, dashboard_template, project_info):
    m = Master(costs_masters, project_info)
    wb = financial_dashboard(m, dashboard_template)
    wb.save("resources/test_dashboards_master_altered.xlsx")


def test_schedule_dashboard(milestone_masters, dashboard_template, project_info):
    m = Master(milestone_masters, project_info)
    milestones = MilestoneData(m, baseline=["all"])
    milestones.filter_chart_info(milestone_type=["Approval", "Delivery"])
    wb = schedule_dashboard(m, milestones, dashboard_template)
    wb.save("resources/test_dashboards_master_altered.xlsx")


def test_benefits_dashboard(benefits_masters, dashboard_template, project_info):
    m = Master(benefits_masters, project_info)
    wb = benefits_dashboard(m, dashboard_template)
    wb.save("resources/test_dashboards_master_altered.xlsx")


def test_overall_dashboard(two_masters, dashboard_template, project_info):
    m = Master(two_masters, project_info)
    milestones = MilestoneData(m, baseline=["all"])
    # milestones.get_milestones(baseline=[])
    wb = overall_dashboard(m, milestones, dashboard_template)
    wb.save("resources/test_dashboards_master_altered.xlsx")

## superceded by test below
# def test_dandelion(basic_masters_dicts, project_info, word_doc):
#     m = Master(basic_masters_dicts, project_info)
#     dand = DandelionData(m, quarter=["standard"])
#     wb = dandelion_data_into_wb(dand)
#     wb.save("resources/test_dandelion_data.xlsx")
#     graph = run_dandelion_matplotlib_chart(dand)
#     put_matplotlib_fig_into_word(word_doc, graph, size=4, transparent=False)
#     word_doc.save("resources/test_dandelion_output.docx")


def test_build_dandelion_graph_auto(basic_masters_dicts, project_info, word_doc):
    m = Master(basic_masters_dicts, project_info)
    d_data = DandelionData(m, quarter=["Q4 18/19"], group=["HSMRPG", "Rail", "RPE"])
    graph = make_a_dandelion_auto(d_data)
    put_matplotlib_fig_into_word(word_doc, graph, size=4, transparent=False)
    word_doc.save("resources/test_dandelion_output.docx")


def test_data_queries_non_milestone(basic_masters_dicts, project_info):
    m = Master(basic_masters_dicts, project_info)
    wb = data_query_into_wb(
        m, keys=["Total Forecast"], quarter=["Q4 18/19", "Q4 17/18", "Q4 16/17"]
    )
    wb.save("resources/test_data_query.xlsx")


def test_data_queries_milestones(milestone_masters, project_info):
    m = Master(milestone_masters, project_info)
    wb = data_query_into_wb(
        m, keys=["Full Operations"], quarter=["Q4 19/20", "Q4 18/19"]
    )
    wb.save("resources/test_data_query_milestones.xlsx")


def test_open_csv_file(key_file):
    l = get_data_query_key_names(key_file)
    assert isinstance(l, (list,))


def test_cal_group_including_removing(milestone_masters, project_info):
    m = Master(milestone_masters, project_info)
    kwargs = {"baseline": "current", "remove": ["Mars"]}
    group = get_group(m, "current", kwargs)
    assert group == ['Sea of Tranquility', 'Apollo 11', 'Apollo 13', 'Falcon 9', 'Columbia']


def test_build_dandelion_graph_manual(build_dandelion, word_doc_landscape):
    dlion = make_a_dandelion_manual(build_dandelion)
    put_matplotlib_fig_into_word(word_doc_landscape, dlion, size=7.5)
    word_doc_landscape.save("resources/dlion_mpl.docx")


































