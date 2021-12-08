"""
Tests for analysis_engine
"""
import configparser
import csv
import os
import datetime
import pickle

import numpy as np
from matplotlib import pyplot as plt
import pytest
from datamaps.api import project_data_from_master

from analysis_engine.ar_data import get_ar_data, ar_run_p_reports
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
    # remove_project_name_from_milestone_key,
    get_sp_data,
    cost_stackplot_graph,
    get_group,
    make_a_dandelion_auto,
    get_horizontal_bar_chart_data,
    simple_horz_bar_chart,
    so_matplotlib,
    radar_chart,
    get_strategic_priorities_data,
    JsonData,
    open_json_file,
    get_project_information, run_p_reports, JsonMaster, build_speedials,
)
from analysis_engine.top35_data import top35_run_p_reports

SOT = "Sea of Tranquility"
A11 = "Apollo 11"
A13 = "Apollo 13"
F9 = "Falcon 9"
COLUMBIA = "Columbia"
MARS = "Mars"
TEST_GROUP = [SOT, A13, F9, COLUMBIA, MARS]


@pytest.mark.skip(reason="old. Pickle not in use.")
def test_master_in_a_pickle(full_test_masters_dict, project_info):
    master = Master(full_test_masters_dict, project_info)
    path_str = str("{0}/resources/test_master".format(os.path.join(os.getcwd())))
    mickle = Pickle(master, path_str)
    assert str(mickle.master.master_data[0].quarter) == "Q1 20/21"


@pytest.mark.slow(reason="lengthy. Only required to inspect JsonData Class.")
def test_json_master_save(full_test_masters_dict, project_info, master_json_path):
    master = JsonMaster(full_test_masters_dict, project_info)
    master.get_baseline_data()
    JsonData(master, master_json_path)


def test_json_master_open(master_json_path):
    jm = open_json_file(master_json_path + ".json")  # jm is json_master
    m = Master(jm)
    assert isinstance(m.master_data, (list,))


@pytest.mark.skip(reason="Old. Pickle not used")
def test_opening_a_pickle(master_pickle_file_path):
    mickle = open_pickle_file(master_pickle_file_path)
    assert str(mickle.master_data[0].quarter) == "Q1 20/21"


def test_creation_of_masters_class(basic_masters_dicts, project_info):
    master = JsonMaster(basic_masters_dicts, project_info)
    assert isinstance(master.master_data, (list,))


def test_creation_of_top250_master_json_file(
        top35_master, top35_project_info, top35_master_json_path
):
    m = JsonMaster(top35_master, top35_project_info, data_type="top35")
    JsonData(m, top35_master_json_path)


def test_getting_baseline_data_from_masters(basic_masters_dicts, project_info):
    master = JsonMaster(basic_masters_dicts, project_info)
    master.get_baseline_data()
    assert isinstance(master.bl_index, (dict,))
    assert master.bl_index["ipdc_milestones"]["Sea of Tranquility"] == [0, 1]
    assert master.bl_index["ipdc_costs"]["Apollo 11"] == [0, 1, 0, 2]
    assert master.bl_index["ipdc_costs"]["Columbia"] == [0, 1, 0, 2]


def test_get_current_project_names(basic_masters_dicts, project_info):
    master = JsonMaster(basic_masters_dicts, project_info)
    assert master.current_projects == [
        "Sea of Tranquility",
        "Apollo 11",
        "Apollo 13",
        "Falcon 9",
        "Columbia",
    ]


def test_getting_project_groups(project_info, basic_masters_dicts):
    m = JsonMaster(basic_masters_dicts, project_info)
    assert isinstance(m.project_stage, (dict,))
    assert isinstance(m.dft_groups, (dict,))


def test_get_project_abbreviations(basic_masters_dicts, project_info):
    master = JsonMaster(basic_masters_dicts, project_info)
    assert master.abbreviations == {
        "Apollo 11": {"abb": "A11", "full name": "Apollo 11"},
        "Apollo 13": {"abb": "A13", "full name": "Apollo 13"},
        "Columbia": {"abb": "Columbia", "full name": "Columbia"},
        "Falcon 9": {"abb": "F9", "full name": "Falcon 9"},
        "Mars": {"abb": "Mars", "full name": "Mars"},
        "Pipe Dreaming": {"abb": "Pdream", "full name": "Pipe Dreaming"},
        "Piping Hot": {"abb": "PH", "full name": "Piping Hot"},
        "Put That in Your Pipe": {"abb": "PtiYP", "full name": "Put That in Your Pipe"},
        "Sea of Tranquility": {"abb": "SoT", "full name": "Sea of Tranquility"},
    }


def test_calculating_spent(master):
    test_dict = master.master_data[0]["data"]
    spent = spent_calculation(test_dict, "Sea of Tranquility")
    assert spent == 1409.33


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


def test_word_doc_heading(word_doc, master):
    wd_heading(word_doc, master, "Apollo 11")
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_contacts(word_doc, master):
    key_contacts(word_doc, master, "Apollo 13")
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_dca_table(word_doc, master):
    dca_table(word_doc, master, "Falcon 9")
    word_doc.save("resources/summary_temp_altered.docx")


def test_word_doc_dca_narratives(word_doc, master):
    dca_narratives(word_doc, master, "Falcon 9")
    word_doc.save("resources/summary_temp_altered.docx")


def test_project_report_meta_data(word_doc, master):
    project = [F9]
    cost = CostData(master, quarter=["standard"], group=project)
    milestones = MilestoneData(master, quarter=["standard"], group=project)
    benefits = BenefitsData(master, quarter=["standard"], group=project)
    project_report_meta_data(word_doc, cost, milestones, benefits, *project)
    word_doc.save("resources/summary_temp_altered.docx")


def test_project_cost_profile_chart(master):
    costs = CostData(master, group=TEST_GROUP, baseline=["standard"])
    costs.get_cost_profile()
    cost_profile_graph(costs, master, chart=False)


def test_project_cost_profile_into_wb(master):
    costs = CostData(master, baseline=["standard"], group=TEST_GROUP)
    costs.get_cost_profile()
    wb = cost_profile_into_wb(costs)
    wb.save("resources/test_cost_profile_output.xlsx")


def test_matplotlib_chart_into_word(word_doc, master):
    costs = CostData(master, group=[F9], baseline=["standard"])
    costs.get_cost_profile()
    graph = cost_profile_graph(costs, master, chart=False)
    change_word_doc_landscape(word_doc)
    put_matplotlib_fig_into_word(word_doc, graph)
    word_doc.save("resources/summary_temp_altered.docx")


def test_get_project_total_costs_benefits_bar_chart(master):
    costs = CostData(master, baseline=["standard"], group=TEST_GROUP)
    benefits = BenefitsData(master, baseline=["standard"], group=TEST_GROUP)
    total_costs_benefits_bar_chart(costs, benefits, master, chart=False)


def test_changing_word_doc_to_landscape(word_doc):
    change_word_doc_landscape(word_doc)
    word_doc.save("resources/summary_changed_to_landscape.docx")


def test_get_stackplot_costs_chart(master):
    sp = get_sp_data(master, group=TEST_GROUP, quarter=["standard"])
    cost_stackplot_graph(sp, master, chart=False)


def test_get_project_total_cost_calculations_for_project(master):
    costs = CostData(master, group=[F9], baseline=["standard"])
    assert costs.c_totals["current"]["spent"] == 471
    assert costs.c_totals["current"]["prof"] == 6281
    assert costs.c_totals["current"]["unprof"] == 0


def test_get_group_total_cost_calculations(master):
    costs = CostData(
        master, group=master.current_projects, quarter=["standard"]
    )
    assert costs.c_totals["Q1 20/21"]["spent"] == 3926
    assert costs.c_totals["Q4 19/20"]["spent"] == 2610


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


def test_getting_benefits_profile_for_a_group(master):
    ben = BenefitsData(
        master, group=master.current_projects, quarter=["standard"]
    )
    assert ben.b_totals["Q1 20/21"]["delivered"] == 0
    assert ben.b_totals["Q1 20/21"]["prof"] == 43659
    assert ben.b_totals["Q1 20/21"]["unprof"] == 10164


def test_getting_benefits_profile_for_a_project(master):
    ben = BenefitsData(master, group=[F9], baseline=["all"])
    assert ben.b_totals["current"]["cat_prof"] == [0, 0, 0, -200]


def test_compare_changes_between_masters(basic_masters_file_paths, project_info):
    gmpp_list = get_gmpp_projects(project_info)
    wb = compare_masters(basic_masters_file_paths, gmpp_list)
    wb.save(os.path.join(os.getcwd(), "resources/cut_down_master_compared.xlsx"))


def test_get_milestone_data_bl(master):
    milestones = MilestoneData(
        master, group=master.current_projects, baseline=["all"]
    )
    assert isinstance(milestones.milestone_dict["current"], (dict,))


def test_get_milestone_data_all(master):
    milestones = MilestoneData(
        master,
        group=master.current_projects,
        quarter=["Q4 19/20", "Q4 18/19"],
    )
    assert isinstance(milestones.milestone_dict["Q4 19/20"], (dict,))


def test_get_milestone_chart_data(master):
    milestones = MilestoneData(master, group=[SOT, A13], baseline=["standard"])
    assert (
        len(milestones.sorted_milestone_dict[milestones.iter_list[0]]["g_dates"]) == 76
    )
    assert (
        len(milestones.sorted_milestone_dict[milestones.iter_list[1]]["g_dates"]) == 76
    )
    assert (
        len(milestones.sorted_milestone_dict[milestones.iter_list[2]]["g_dates"]) == 76
    )


def test_compile_milestone_chart_with_filter(master):
    milestones = MilestoneData(
        master, group=[SOT], quarter=["Q4 19/20", "Q4 18/19"]
    )
    milestones.filter_chart_info(dates=["1/1/2013", "1/1/2014"])
    milestone_chart(
        milestones,
        master,
        title="Group Test",
        blue_line="Today",
        chart=False,
    )


def test_removing_project_name_from_milestone_keys(master):
    """
    The standard list contained with in the sorted_milestone_dict is {"names": ["Project Name,
    Milestone Name, ...]. When there is only one project in the dictionary the need for a Project
    Name is obsolete. The function remove_project_name_from_milestone_key, removes the project name
    and returns milestone name only.
    """
    milestones = MilestoneData(master, group=[SOT], baseline=["all"])
    milestones.filter_chart_info(dates=["1/1/2013", "1/1/2014"])
    key_names = milestones.sorted_milestone_dict["current"]["names"]
    # key_names = remove_project_name_from_milestone_key("SoT", key_names)
    assert key_names == [
        "Sputnik Radiation",
        "Lunar Magma",
        "Standard B",
        "Standard C",
        "Mercury Eleven",
        "Tranquility Radiation",
    ]


def test_putting_milestones_into_wb(master):
    milestones = MilestoneData(master, group=[SOT], baseline=["all"])
    milestones.filter_chart_info(dates=["1/1/2013", "1/1/2014"])
    wb = put_milestones_into_wb(milestones)
    wb.save("resources/milestone_data_output_test.xlsx")


def test_dca_analysis(master):
    dca = DcaData(master, quarter=["standard"])
    wb = dca_changes_into_excel(dca)
    wb.save("resources/dca_print.xlsx")


def test_speedial_print_out(master, word_doc):
    dca = DcaData(master, quarter=["standard"], conf_type='sro')
    dca.get_changes()
    dca_changes_into_word(dca, word_doc)
    word_doc.save("resources/dca_checks.docx")


def test_speedial_graph(master, word_doc):
    dca_data = DcaData(master, quarter=["standard"], conf_type='sro', rag_number="3")
    dca_data.get_changes()
    build_speedials(dca_data, word_doc)
    word_doc.save("resources/speedial_graph.docx")


def test_risk_analysis(master):
    risk = RiskData(master, quarter=["standard"])
    wb = risks_into_excel(risk)
    wb.save("resources/risks.xlsx")


def test_vfm_analysis(master):
    vfm = VfMData(master, quarter=["standard"])
    wb = vfm_into_excel(vfm)
    wb.save("resources/vfm.xlsx")


def test_sorting_project_by_dca(master):
    rag_list = sort_projects_by_dca(master.master_data[0], TEST_GROUP)
    assert rag_list == [
        ("Falcon 9", "Amber"),
        ("Sea of Tranquility", "Amber/Green"),
        ("Apollo 13", "Amber/Green"),
        ("Mars", "Amber/Green"),
        ("Columbia", "Green"),
    ]


@pytest.mark.skip(reason="failing need to look at.")
def test_calculating_wlc_changes(master):
    costs = CostData(
        master, group=[master.current_projects], baseline=["all"]
    )
    costs.calculate_wlc_change()
    assert costs.wlc_change == {
        "Apollo 13": {"baseline one": 0, "last quarter": 0},
        "Columbia": {"baseline one": -43, "last quarter": -43},
        "Falcon 9": {"baseline one": 5, "last quarter": 5},
        "Mars": {"baseline one": 0},
        "Sea of Tranquility": {"baseline one": 54, "last quarter": 54},
    }


@pytest.mark.skip(reason="passing but empty dict so not right.")
def test_calculating_schedule_changes(master):
    milestones = MilestoneData(master, group=[SOT, A11, A13])
    milestones.calculate_schedule_changes()
    assert isinstance(milestones.schedule_change, (dict,))


def test_printout_of_milestones(word_doc, master):
    milestones = MilestoneData(master, group=[SOT], baseline=["standard"])
    change_word_doc_landscape(word_doc)
    print_out_project_milestones(word_doc, milestones, SOT)
    word_doc.save("resources/summary_temp_altered.docx")


@pytest.mark.skip(reason="failing need to look at.")
def test_cost_schedule_matrix(master, project_info):
    costs = CostData(
        master, group=master.current_projects, quarters=["standard"]
    )
    milestones = MilestoneData(master, group=master.current_projects)
    milestones.calculate_schedule_changes()
    wb = cost_v_schedule_chart_into_wb(milestones, costs)
    wb.save("resources/test_costs_schedule_matrix.xlsx")


def test_financial_dashboard(master, dashboard_template):
    wb = financial_dashboard(master, dashboard_template)
    wb.save("resources/test_dashboards_master_altered.xlsx")


def test_schedule_dashboard(master, dashboard_template):
    milestones = MilestoneData(master, baseline=["all"], group=[master.current_projects])
    milestones.filter_chart_info(milestone_type=["Approval", "Delivery"])
    wb = schedule_dashboard(master, milestones, dashboard_template)
    wb.save("resources/test_dashboards_master_altered.xlsx")


def test_benefits_dashboard(master, dashboard_template):
    wb = benefits_dashboard(master, dashboard_template)
    wb.save("resources/test_dashboards_master_altered.xlsx")


@pytest.mark.skip(reason="need to reconfigure test so it's correct.")
def test_overall_dashboard(master, dashboard_template):
    milestones = MilestoneData(master, baseline=["all"])
    wb = overall_dashboard(master, milestones, dashboard_template)
    wb.save("resources/test_dashboards_master_altered.xlsx")


def test_build_dandelion_graph(word_doc_landscape, ipdc_data, master):
    # m = Master(*d["data"], **d["op_args"])  # currently necessary for cdg and top35 data
    # dl_data = DandelionData(m, **d["op_args"])
    dl_data = DandelionData(master, **ipdc_data["op_args"])
    d_lion = make_a_dandelion_auto(dl_data, **ipdc_data["op_args"])
    put_matplotlib_fig_into_word(word_doc_landscape, d_lion, size=7)
    word_doc_landscape.save(ipdc_data["docx_save_path"].format("ipdc_d_graph"))


def test_data_queries_non_milestone(master):
    wb = data_query_into_wb(
        master,
        key=["Total Forecast"],
        quarter=["Q4 18/19", "Q4 17/18", "Q4 16/17"],
        group=[A11],
    )
    wb.save("resources/test_data_query.xlsx")


def test_data_queries_milestones(master):
    wb = data_query_into_wb(
        master,
        key=["Full Operations"],
        quarter=["Q4 19/20", "Q4 18/19"],
        group=[SOT],
    )
    wb.save("resources/test_data_query_milestones.xlsx")


def test_open_csv_file(key_file):
    key_list = get_data_query_key_names(key_file)
    assert isinstance(key_list, (list,))


@pytest.mark.skip(reason="Failing. get_group function messy and could use refactor.")
def test_cal_group_including_removing(master):
    op_args = {"baseline": "standard", "remove": ["Mars"]}
    group = get_group(master, "Q1 20/21", **op_args)
    assert group == [
        "Sea of Tranquility",
        "Apollo 11",
        "Apollo 13",
        "Falcon 9",
        "Columbia",
    ]


@pytest.mark.skip(reason="not currently in use.")
def test_build_dandelion_graph_manual(build_dandelion, word_doc_landscape):
    dlion = make_a_dandelion_manual(build_dandelion)
    put_matplotlib_fig_into_word(word_doc_landscape, dlion, size=7.5)
    word_doc_landscape.save("resources/dlion_mpl.docx")


@pytest.mark.skip(reason="wp")
def test_build_horizontal_bar_chart_manually(
    horizontal_bar_chart_data, word_doc_landscape
):
    # graph = get_horizontal_bar_chart_data(horizontal_bar_chart_data)
    simple_horz_bar_chart(horizontal_bar_chart_data)
    # put_matplotlib_fig_into_word(word_doc_landscape, graph)
    # word_doc_landscape.save("resources/distributed_horz_bar_chart.docx")
    # so_matplotlib()


def test_radar_chart(sp_data, master, word_doc):
    chart = radar_chart(sp_data, master, chart=False)
    put_matplotlib_fig_into_word(word_doc, chart, size=5)
    word_doc.save("resources/test_radar.docx")


def test_strategic_priority_data(sp_data, master):
    sp_dict = get_strategic_priorities_data(sp_data, master)
    assert isinstance(sp_dict, (list,))


@pytest.mark.skip(reason="temp code for now. No plans for long term ae intergration")
def test_annual_report_summaries():
    data = get_ar_data()
    pi = get_project_information()
    ar_run_p_reports(data, pi)


def test_top35_summaries(top35_data):
    top35_run_p_reports(top35_data["master"], **top35_data["op_args"])


def test_match_data_types():
    dft_val = 1664.71708896933
    gmpp_val = 1665
    if isinstance(dft_val, float) and isinstance(gmpp_val, int):
        dft_val = round(dft_val)
        # gmpp_val = int(gmpp_val)

    assert dft_val == 1665
    assert gmpp_val == 1665


def test_calculate_group_angles_dandelion():
    group_five = ["HSRG", "RSS", "SAUSAGE", "BACON", "EGGS"]
    group_four = ["BEANS", "BACON", "EGGS", "TOAST"]
    group_three = ["SAUSAGE", "BACON", "EGGS"]
    group_two = ["BACON", "EGGS"]
    group = group_four
 
    # Dandelion graph needs an algorithm to calculate the distribution
    # of group circles. The circles are placed and distributed left
    # to right around the center circle. 
    angle_list = []
    # start_point needs to come down as numbers increase
    start_point = 290 * ((29 - ((len(group))-2)) / 29)
    # distribution increase needs to come down as numbers increase
    distribution_start = 0
    distribution_increase = 140
    if len(group) > 2:  # no change in distribution increase if group of two
        for i in range(len(group)):
            distribution_increase = distribution_increase*0.82
    for i in range(len(group)):
        angle = distribution_start + start_point
        if angle > 360:
            angle = angle - 360
        angle_list.append(int(angle))
        distribution_start += distribution_increase
    assert isinstance(angle_list, (list,))

