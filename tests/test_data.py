import click
from docx.shared import Inches

from analysis_engine.error_msgs import ProjectNameError, ProjectGroupError, ProjectStageError, logger, InputError
from analysis_engine.gmpp_int import GmppOnlineCosts
from analysis_engine.merge import Merge
from analysis_engine.risks import RiskData, portfolio_risks_into_excel, risks_into_excel, \
    portfolio_risks_into_word_by_project, portfolio_risks_into_word_by_risk
from analysis_engine.summaries import run_p_reports
from tests.test_op_args import *

from analysis_engine.main import CliOpArgs
from analysis_engine.settings import return_koi_fn_keys, get_masters_to_merge
from analysis_engine.core_data import (
    PythonMasterData,
    get_master_data,
    get_group_meta_data,
    get_stage_meta_data,
    get_project_information,
    JsonData,
    open_json_file, report_config, get_dandelion_meta_data,
)

from analysis_engine.dandelion import DandelionData, make_a_dandelion_auto
from analysis_engine.dca import DcaData, dca_changes_into_word, dca_changes_into_excel
from analysis_engine.speed_dials import build_speed_dials
from analysis_engine.render_utils import put_matplotlib_fig_into_word, get_input_doc
from analysis_engine.dashboards import narrative_dashboard, cdg_dashboard, ipdc_dashboard
from analysis_engine.milestones import (
    MilestoneData,
    milestone_chart,
    put_milestones_into_wb,
)
from analysis_engine.query import data_query_into_wb
from analysis_engine.costs import CostData, cost_profile_graph_new, cost_profile_into_wb_new

SETTINGS_DICT = report_config(REPORTING_TYPE)


def test_get_project_information():
    project_info = get_project_information(SETTINGS_DICT)
    assert isinstance(project_info, dict)


def test_get_group_metadata_from_config():
    GROUP_META = get_group_meta_data(SETTINGS_DICT)
    STAGE_META = get_stage_meta_data(SETTINGS_DICT)

    META = {**GROUP_META, **STAGE_META}

    assert isinstance(META, dict)


def test_get_raw_master_data_in_list():
    md = get_master_data(SETTINGS_DICT)
    assert isinstance(md, list)


def test_saving_creating_json_master():
    GROUP_META = get_group_meta_data(SETTINGS_DICT)
    STAGE_META = get_stage_meta_data(SETTINGS_DICT)
    PORT_META = get_dandelion_meta_data(SETTINGS_DICT)
    META = {**GROUP_META, **STAGE_META, **PORT_META}

    try:
        master = PythonMasterData(
            get_master_data(SETTINGS_DICT),
            get_project_information(SETTINGS_DICT),
            META,
            data_type=SETTINGS_DICT["report"],
        )
    except (ProjectNameError, ProjectGroupError, ProjectStageError) as e:
        logger.critical(e)
        pass

    master_json_path = str(
        "{0}/core_data/json/master".format(SETTINGS_DICT["root_path"])
    )
    JsonData(master, master_json_path)


def test_json_master():
    data = open_json_file(
        f"/home/will/Documents/{SETTINGS_DICT['report']}/core_data/json/master.json"
    )
    assert isinstance(data["master_data"], (list,))


def test_get_project_abbreviations():
    data = open_json_file(
        f"/home/will/Documents/{REPORTING_TYPE}/core_data/json/master.json"
    )
    abb_list = []
    for p in data['project_information']:
        abb_list.append(data['project_information'][p]['Abbreviations'])

    assert len(abb_list) != 0


def test_build_dandelion_graph():
    for x in DANDELION_OP_ARGS_DICT:
        print(x['test_name'])
        cli = CliOpArgs(x, SETTINGS_DICT)
        if cli.combined_args['report'] == 'ipdc':
            cli.combined_args['abbreviations'] = True

        d_data = DandelionData(cli.md, **cli.combined_args)
        if 'env_funds' in cli.combined_args:
            if not cli.combined_args['env_funds']:  # empty list
                d_data.get_environmental_fund_data()

        if cli.combined_args["chart"] != "save":
            make_a_dandelion_auto(d_data, **cli.combined_args)
        else:
            d_graph = make_a_dandelion_auto(d_data, **cli.combined_args)
            doc_path = (
                    str(cli.combined_args["root_path"]) + cli.combined_args["word_landscape"]
            )
            doc = get_input_doc(doc_path)
            put_matplotlib_fig_into_word(doc, d_graph, width=Inches(8))
            doc_output_path = (
                    str(cli.combined_args["root_path"]) + cli.combined_args["word_save_path"]
            )
            doc.save(doc_output_path.format(f"{x['test_name']}"))


def test_dca_analysis():
    for x in SPEED_DIAL_AND_DCA_OP_ARGS:
        print(x['test_name'])
        x["subparser_name"] = "dcas"
        try:
            cli = CliOpArgs(x, SETTINGS_DICT)
            sdmd = DcaData(cli.md, **cli.combined_args)
            sdmd.get_changes()
            changes_doc = dca_changes_into_word(
                sdmd, str(cli.combined_args["root_path"]) + cli.combined_args["word_portrait"]
            )
            changes_doc.save(
                str(cli.combined_args["root_path"])
                + cli.settings["word_save_path"].format(f"dca_changes_{x['test_name']}")
            )
            wb = dca_changes_into_excel(sdmd)
            wb.save(cli.combined_args["root_path"]
                    + cli.settings["excel_save_path"].format(f"dca_changes_{x['test_name']}")
                    )
        except InputError as e:
            logger.critical(e)
            pass


def test_speed_dials():
    for x in SPEED_DIAL_AND_DCA_OP_ARGS:
        print(x['test_name'])
        cli = CliOpArgs(x, SETTINGS_DICT)

        if cli.combined_args['report'] == 'ipdc':
            cli.combined_args["rag_number"] = '3'
        if cli.combined_args['report'] == 'cdg':
            cli.combined_args["rag_number"] = '5'

        try:
            sdmd = DcaData(cli.md, **cli.combined_args)
            sdmd.get_changes()
            sd_doc = get_input_doc(
                str(cli.combined_args["root_path"]) + cli.combined_args["word_landscape"]
            )
            build_speed_dials(sdmd, sd_doc)
            sd_doc.save(
                str(cli.combined_args["root_path"])
                + cli.settings["word_save_path"].format(f"speed_dials_{x['test_name']}")
            )
        except InputError as e:
            logger.critical(e)
            pass


def test_dashboards():
    op_args = {'subparser_name': 'dashboards'}
    cli = CliOpArgs(op_args, SETTINGS_DICT)
    if REPORTING_TYPE == 'cdg':
        narrative_d_master = get_input_doc(
            str(cli.combined_args["root_path"]) + cli.combined_args["narrative_dashboard"]
        )
        narrative_dashboard(cli.md, narrative_d_master)  #
        narrative_d_master.save(
            str(cli.combined_args["root_path"])
            + cli.combined_args["excel_save_path"].format("narrative_dashboard_completed")
        )
        cdg_d_master = get_input_doc(
            str(cli.combined_args["root_path"]) + cli.combined_args["dashboard"]
        )
        cdg_dashboard(cli.md, cdg_d_master)
        cdg_d_master.save(
            str(cli.combined_args["root_path"])
            + cli.combined_args["excel_save_path"].format("dashboard_completed")
        )
    if REPORTING_TYPE == 'ipdc':
        ipdc_d_master = get_input_doc(
            str(cli.combined_args["root_path"]) + cli.combined_args["dashboard"]
        )
        ipdc_dashboard(cli.md, ipdc_d_master, **cli.combined_args)
        ipdc_d_master.save(
            str(cli.combined_args["root_path"])
            + cli.combined_args["excel_save_path"].format("dashboards_completed")
        )


def test_milestones():
    for x in MILESTONES_OP_ARGS:
        print(x['test_name'])
        cli = CliOpArgs(x, SETTINGS_DICT)
        ms = MilestoneData(cli.md, **cli.combined_args)
        if (
                # "type" in combined_args  # NOT IN USE.
                "dates" in cli.combined_args
                or "koi" in cli.combined_args
                or "koi_fn" in cli.combined_args
        ):
            return_koi_fn_keys(cli.combined_args)
            ms.filter_chart_info(**cli.combined_args)

        if cli.combined_args["chart"] == "save":
            ms_graph = milestone_chart(ms, **cli.combined_args)
            doc = get_input_doc(
                str(cli.combined_args["root_path"]) + cli.combined_args["word_landscape"]
            )
            put_matplotlib_fig_into_word(doc, ms_graph, width=Inches(8))
            doc.save(
                str(cli.combined_args["root_path"])
                + cli.combined_args["word_save_path"].format(f"milestones_{x['test_name']}")
            )
        if cli.combined_args["chart"] == "show":
            milestone_chart(ms, **cli.combined_args)

        wb = put_milestones_into_wb(ms)
        wb.save(cli.combined_args["root_path"]
            + cli.settings["excel_save_path"].format(f"milestones_{x['test_name']}")
        )


def test_query():
    for x in QUERY_ARGS:
        print(x['test_name'])
        cli = CliOpArgs(x, SETTINGS_DICT)
        op_args = return_koi_fn_keys(cli.combined_args)
        wb = data_query_into_wb(cli.md, **op_args)
        wb.save(
            str(cli.settings["root_path"])
            + cli.settings["excel_save_path"].format(f"{x['test_name']}")
        )


def test_gmpp_online_data():
    if REPORTING_TYPE == 'ipdc':
        cli = CliOpArgs({'subparser_name': 'gmpp_data'}, SETTINGS_DICT)
        md = GmppOnlineCosts(**cli.combined_args)
        md.place_into_dft_master_format()
        # md.put_cost_totals_into_wb()


def test_merge_masters():
    if REPORTING_TYPE == 'ipdc':
        cli = CliOpArgs({'subparser_name': 'merge_masters'}, SETTINGS_DICT)
        get_masters_to_merge(cli.combined_args)
        Merge(**cli.combined_args)


def test_portfolio_risks_excel():
    if REPORTING_TYPE == 'ipdc':
        for x in PORT_RISK_OP_ARGS:
            print(x['test_name'])
            cli = CliOpArgs(x, SETTINGS_DICT)
            rd = RiskData(cli.md, **cli.combined_args)
            wb = portfolio_risks_into_excel(rd)
            wb.save(
                str(cli.settings["root_path"])
                + cli.settings["excel_save_path"].format(f"{x['test_name']}")
            )


def test_risks_excel():
    if REPORTING_TYPE == 'ipdc':
        for x in PORT_RISK_OP_ARGS:
            print(x['test_name'] + '_RISKS')
            x["subparser_name"] = "risks_project"
            cli = CliOpArgs(x, SETTINGS_DICT)
            rd = RiskData(cli.md, **cli.combined_args)
            wb = risks_into_excel(rd)
            wb.save(
                str(cli.settings["root_path"])
                + cli.settings["excel_save_path"].format(f"{x['test_name']}_Risks")
            )


def test_risks_word():
    if REPORTING_TYPE == 'ipdc':
        for x in PORT_RISK_OP_WORD_ARGS:
            print(x['test_name'] + '_RISKS_WORD')
            x["subparser_name"] = "risks_printout"
            cli = CliOpArgs(x, SETTINGS_DICT)
            rd = RiskData(cli.md, **cli.combined_args)
            by_proj_doc = portfolio_risks_into_word_by_project(rd)
            by_proj_doc.save(
                str(cli.settings["root_path"]) + cli.settings["word_save_path"].format(
                    f"{x['test_name']}_risks_printout_by_project")
            )
            by_risk_doc = portfolio_risks_into_word_by_risk(rd)
            by_risk_doc.save(
                str(cli.settings["root_path"]) + cli.settings["word_save_path"].format(
                    f"{x['test_name']}_risks_printout_by_risk")
            )


def test_costs():
    if REPORTING_TYPE == 'ipdc':
        for x in COST_OP_ARGS:
            print(x['test_name'])
            cli = CliOpArgs(x, SETTINGS_DICT)
            c = CostData(cli.md, **cli.combined_args)
            c.get_forecast_cost_profile()
            wb = cost_profile_into_wb_new(c)
            if cli.combined_args["chart"] == "save":
                ms_graph = cost_profile_graph_new(c, **cli.combined_args)
                doc = get_input_doc(
                    str(cli.combined_args["root_path"]) + cli.combined_args["word_landscape"]
                )
                put_matplotlib_fig_into_word(doc, ms_graph, width=Inches(8))
                doc.save(
                    str(cli.combined_args["root_path"])
                    + cli.combined_args["word_save_path"].format(f"{x['test_name']}")
                )
            if cli.combined_args["chart"] == "show":
                cost_profile_graph_new(c, **cli.combined_args)

            wb.save(
                str(cli.settings["root_path"])
                + cli.settings["excel_save_path"].format(f"{x['test_name']}")
            )


def test_summaries():
    if REPORTING_TYPE == 'ipdc':
        for x in SUM_OP_ARGS:
            print(x['test_name'])
            cli = CliOpArgs(x, SETTINGS_DICT)
            run_p_reports(cli.md, **cli.combined_args)


# def test_summaries():
#     if programme == "summaries":
#     op_args["quarter"] = [str(m.current_quarter)]
#     if "type" not in op_args:
#         op_args["type"] = "short"
#     run_p_reports(m, **op_args)

# @pytest.mark.skip(reason="refactor required")
# def test_calculating_spent(master):
#     test_dict = master.master_data[0]["data"]
#     spent = spent_calculation(test_dict, "Sea of Tranquility")
#     assert spent == 1409.33
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_open_word_doc(word_doc):
#     word_doc.add_paragraph(
#         "Because i'm still in love with you I want to see you dance again, "
#         "because i'm still in love with you on this harvest moon"
#     )
#     word_doc.save("resources/summary_temp_altered.docx")
#     var = word_doc.paragraphs[1].text
#     assert (
#         "Because i'm still in love with you I want to see you dance again, "
#         "because i'm still in love with you on this harvest moon" == var
#     )
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_word_doc_heading(word_doc, master):
#     wd_heading(word_doc, master, "Apollo 11")
#     word_doc.save("resources/summary_temp_altered.docx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_word_doc_contacts(word_doc, master):
#     key_contacts(word_doc, master, "Apollo 13")
#     word_doc.save("resources/summary_temp_altered.docx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_word_doc_dca_table(word_doc, master):
#     dca_table(word_doc, master, "Falcon 9")
#     word_doc.save("resources/summary_temp_altered.docx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_word_doc_dca_narratives(word_doc, master):
#     dca_narratives(word_doc, master, "Falcon 9")
#     word_doc.save("resources/summary_temp_altered.docx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_project_report_meta_data(word_doc, master):
#     project = [F9]
#     cost = CostData(master, quarter=["standard"], group=project)
#     milestones = MilestoneData(master, quarter=["standard"], group=project)
#     benefits = BenefitsData(master, quarter=["standard"], group=project)
#     project_report_meta_data(word_doc, cost, milestones, benefits, *project)
#     word_doc.save("resources/summary_temp_altered.docx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_project_cost_profile_chart(master):
#     costs = CostData(master, group=TEST_GROUP, baseline=["standard"])
#     costs.get_cost_profile()
#     cost_profile_graph(costs, master, chart=False)
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_project_cost_profile_into_wb(master):
#     costs = CostData(master, baseline=["standard"], group=TEST_GROUP)
#     costs.get_cost_profile()
#     wb = cost_profile_into_wb(costs)
#     wb.save("resources/test_cost_profile_output.xlsx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_matplotlib_chart_into_word(word_doc, master):
#     costs = CostData(master, group=[F9], baseline=["standard"])
#     costs.get_cost_profile()
#     graph = cost_profile_graph(costs, master, chart=False)
#     change_word_doc_landscape(word_doc)
#     put_matplotlib_fig_into_word(word_doc, graph)
#     word_doc.save("resources/summary_temp_altered.docx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_get_project_total_costs_benefits_bar_chart(master):
#     costs = CostData(master, baseline=["standard"], group=TEST_GROUP)
#     benefits = BenefitsData(master, baseline=["standard"], group=TEST_GROUP)
#     total_costs_benefits_bar_chart(costs, benefits, master, chart=False)
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_changing_word_doc_to_landscape(word_doc):
#     change_word_doc_landscape(word_doc)
#     word_doc.save("resources/summary_changed_to_landscape.docx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_get_stackplot_costs_chart(master):
#     sp = get_sp_data(master, group=TEST_GROUP, quarter=["standard"])
#     cost_stackplot_graph(sp, master, chart=False)
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_get_project_total_cost_calculations_for_project(master):
#     costs = CostData(master, group=[F9], baseline=["standard"])
#     assert costs.totals["current"]["spent"] == 471
#     assert costs.totals["current"]["prof"] == 6281
#     assert costs.totals["current"]["unprof"] == 0
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_get_group_total_cost_calculations(master):
#     costs = CostData(master, group=master.current_projects, quarter=["standard"])
#     assert costs.totals["Q1 20/21"]["spent"] == 3926
#     assert costs.totals["Q4 19/20"]["spent"] == 2610
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_put_change_keys_into_a_dict(change_log):
#     keys_dict = put_key_change_master_into_dict(change_log)
#     assert isinstance(keys_dict, (dict,))
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_altering_master_wb_file_key_names(change_log, list_cost_masters_files):
#     keys_dict = put_key_change_master_into_dict(change_log)
#     run_change_keys(list_cost_masters_files, keys_dict)
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_getting_benefits_profile_for_a_group(master):
#     ben = BenefitsData(master, group=master.current_projects, quarter=["standard"])
#     assert ben.b_totals["Q1 20/21"]["delivered"] == 0
#     assert ben.b_totals["Q1 20/21"]["prof"] == 43659
#     assert ben.b_totals["Q1 20/21"]["unprof"] == 10164
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_getting_benefits_profile_for_a_project(master):
#     ben = BenefitsData(master, group=[F9], baseline=["all"])
#     assert ben.b_totals["current"]["cat_prof"] == [0, 0, 0, -200]
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_compare_changes_between_masters(basic_masters_file_paths, project_info):
#     gmpp_list = get_gmpp_projects(project_info)
#     wb = compare_masters(basic_masters_file_paths, gmpp_list)
#     wb.save(os.path.join(os.getcwd(), "resources/cut_down_master_compared.xlsx"))
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_get_milestone_data_bl(master):
#     milestones = MilestoneData(master, group=master.current_projects, baseline=["all"])
#     assert isinstance(milestones.milestone_dict["current"], (dict,))
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_get_milestone_data_all(master):
#     milestones = MilestoneData(
#         master,
#         group=master.current_projects,
#         quarter=["Q4 19/20", "Q4 18/19"],
#     )
#     assert isinstance(milestones.milestone_dict["Q4 19/20"], (dict,))
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_get_milestone_chart_data(master):
#     milestones = MilestoneData(master, group=[SOT, A13], baseline=["standard"])
#     assert (
#         len(milestones.sorted_milestone_dict[milestones.iter_list[0]]["g_dates"]) == 76
#     )
#     assert (
#         len(milestones.sorted_milestone_dict[milestones.iter_list[1]]["g_dates"]) == 76
#     )
#     assert (
#         len(milestones.sorted_milestone_dict[milestones.iter_list[2]]["g_dates"]) == 76
#     )
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_compile_milestone_chart_with_filter(master):
#     milestones = MilestoneData(master, group=[SOT], quarter=["Q4 19/20", "Q4 18/19"])
#     milestones.filter_chart_info(dates=["1/1/2013", "1/1/2014"])
#     milestone_chart(
#         milestones,
#         master,
#         title="Group Test",
#         blue_line="Today",
#         chart=False,
#     )
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_removing_project_name_from_milestone_keys(master):
#     """
#     The standard list contained with in the sorted_milestone_dict is {"names": ["Project Name,
#     Milestone Name, ...]. When there is only one project in the dictionary the need for a Project
#     Name is obsolete. The function remove_project_name_from_milestone_key, removes the project name
#     and returns milestone name only.
#     """
#     milestones = MilestoneData(master, group=[SOT], baseline=["all"])
#     milestones.filter_chart_info(dates=["1/1/2013", "1/1/2014"])
#     key_names = milestones.sorted_milestone_dict["current"]["names"]
#     # key_names = remove_project_name_from_milestone_key("SoT", key_names)
#     assert key_names == [
#         "Sputnik Radiation",
#         "Lunar Magma",
#         "Standard B",
#         "Standard C",
#         "Mercury Eleven",
#         "Tranquility Radiation",
#     ]
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_putting_milestones_into_wb(master):
#     milestones = MilestoneData(master, group=[SOT], baseline=["all"])
#     milestones.filter_chart_info(dates=["1/1/2013", "1/1/2014"])
#     wb = put_milestones_into_wb(milestones)
#     wb.save("resources/milestone_data_output_test.xlsx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_speedial_print_out(master, word_doc):
#     dca = DcaData(master, quarter=["standard"], conf_type="sro")
#     dca.get_changes()
#     dca_changes_into_word(dca, word_doc)
#     word_doc.save("resources/dca_checks.docx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_risk_analysis(master):
#     risk = RiskData(master, quarter=["standard"])
#     wb = risks_into_excel(risk)
#     wb.save("resources/risks.xlsx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_vfm_analysis(master):
#     vfm = VfMData(master, quarter=["standard"])
#     wb = vfm_into_excel(vfm)
#     wb.save("resources/vfm.xlsx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_sorting_project_by_dca(master):
#     rag_list = sort_projects_by_dca(master.master_data[0], TEST_GROUP)
#     assert rag_list == [
#         ("Falcon 9", "Amber"),
#         ("Sea of Tranquility", "Amber/Green"),
#         ("Apollo 13", "Amber/Green"),
#         ("Mars", "Amber/Green"),
#         ("Columbia", "Green"),
#     ]
#
#
# @pytest.mark.skip(reason="failing need to look at.")
# def test_calculating_wlc_changes(master):
#     costs = CostData(master, group=[master.current_projects], baseline=["all"])
#     costs.calculate_wlc_change()
#     assert costs.wlc_change == {
#         "Apollo 13": {"baseline one": 0, "last quarter": 0},
#         "Columbia": {"baseline one": -43, "last quarter": -43},
#         "Falcon 9": {"baseline one": 5, "last quarter": 5},
#         "Mars": {"baseline one": 0},
#         "Sea of Tranquility": {"baseline one": 54, "last quarter": 54},
#     }
#
#
# @pytest.mark.skip(reason="passing but empty dict so not right.")
# def test_calculating_schedule_changes(master):
#     milestones = MilestoneData(master, group=[SOT, A11, A13])
#     milestones.calculate_schedule_changes()
#     assert isinstance(milestones.schedule_change, (dict,))
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_printout_of_milestones(word_doc, master):
#     milestones = MilestoneData(master, group=[SOT], baseline=["standard"])
#     change_word_doc_landscape(word_doc)
#     print_out_project_milestones(word_doc, milestones, SOT)
#     word_doc.save("resources/summary_temp_altered.docx")
#
#
# @pytest.mark.skip(reason="failing need to look at.")
# def test_cost_schedule_matrix(master, project_info):
#     costs = CostData(master, group=master.current_projects, quarters=["standard"])
#     milestones = MilestoneData(master, group=master.current_projects)
#     milestones.calculate_schedule_changes()
#     wb = cost_v_schedule_chart_into_wb(milestones, costs)
#     wb.save("resources/test_costs_schedule_matrix.xlsx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_financial_dashboard(master, dashboard_template):
#     wb = financial_dashboard(master, dashboard_template)
#     wb.save("resources/test_dashboards_master_altered.xlsx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_schedule_dashboard(master, dashboard_template):
#     milestones = MilestoneData(
#         master, baseline=["all"], group=[master.current_projects]
#     )
#     milestones.filter_chart_info(milestone_type=["Approval", "Delivery"])
#     wb = schedule_dashboard(master, milestones, dashboard_template)
#     wb.save("resources/test_dashboards_master_altered.xlsx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_benefits_dashboard(master, dashboard_template):
#     wb = benefits_dashboard(master, dashboard_template)
#     wb.save("resources/test_dashboards_master_altered.xlsx")
#
#
# @pytest.mark.skip(reason="need to reconfigure test so it's correct.")
# def test_overall_dashboard(master, dashboard_template):
#     milestones = MilestoneData(master, baseline=["all"])
#     wb = overall_dashboard(master, milestones, dashboard_template)
#     wb.save("resources/test_dashboards_master_altered.xlsx")
#
#
# @pytest.mark.skip(reason="refactor required")
# def test_open_csv_file(key_file):
#     key_list = get_data_query_key_names(key_file)
#     assert isinstance(key_list, (list,))
#
#
# @pytest.mark.skip(reason="Failing. get_group function messy and could use refactor.")
# def test_cal_group_including_removing(master):
#     op_args = {"baseline": "standard", "remove": ["Mars"]}
#     group = get_group(master, "Q1 20/21", **op_args)
#     assert group == [
#         "Sea of Tranquility",
#         "Apollo 11",
#         "Apollo 13",
#         "Falcon 9",
#         "Columbia",
#     ]
#
#
# @pytest.mark.skip(reason="not currently in use.")
# def test_build_dandelion_graph_manual(build_dandelion, word_doc_landscape):
#     dlion = make_a_dandelion_manual(build_dandelion)
#     put_matplotlib_fig_into_word(word_doc_landscape, dlion, size=7.5)
#     word_doc_landscape.save("resources/dlion_mpl.docx")
#
#
# @pytest.mark.skip(reason="wp")
# def test_build_horizontal_bar_chart_manually(
#     horizontal_bar_chart_data, word_doc_landscape
# ):
#     # graph = get_horizontal_bar_chart_data(horizontal_bar_chart_data)
#     simple_horz_bar_chart(horizontal_bar_chart_data)
#     # put_matplotlib_fig_into_word(word_doc_landscape, graph)
#     # word_doc_landscape.save("resources/distributed_horz_bar_chart.docx")
#     # so_matplotlib()
#
#
# @pytest.mark.skip(reason="not currently in use.")
# def test_radar_chart(sp_data, master, word_doc):
#     chart = radar_chart(sp_data, master, chart=False)
#     put_matplotlib_fig_into_word(word_doc, chart, size=5)
#     word_doc.save("resources/test_radar.docx")
#
# @pytest.mark.skip(reason="not currently in use.")
# def test_strategic_priority_data(sp_data, master):
#     sp_dict = get_strategic_priorities_data(sp_data, master)
#     assert isinstance(sp_dict, (list,))
#
#
# @pytest.mark.skip(reason="temp code for now. No plans for long term ae intergration")
# def test_annual_report_summaries():
#     data = get_ar_data()
#     pi = get_project_information()
#     ar_run_p_reports(data, pi)
#
#
# def test_top35_summaries(top35_data):
#     top35_run_p_reports(top35_data["master"], **top35_data["op_args"])
#
#
# @pytest.mark.skip(reason="not currently in use.")
# def test_match_data_types():
#     dft_val = 1664.71708896933
#     gmpp_val = 1665
#     if isinstance(dft_val, float) and isinstance(gmpp_val, int):
#         dft_val = round(dft_val)
#         # gmpp_val = int(gmpp_val)
#
#     assert dft_val == 1665
#     assert gmpp_val == 1665
#
#
# @pytest.mark.skip(reason="not currently in use.")
# def test_calculate_group_angles_dandelion():
#     group_five = ["HSRG", "RSS", "SAUSAGE", "BACON", "EGGS"]
#     group_four = ["BEANS", "BACON", "EGGS", "TOAST"]
#     group_three = ["SAUSAGE", "BACON", "EGGS"]
#     group_two = ["BACON", "EGGS"]
#     group = group_four
#
#     # Dandelion graph needs an algorithm to calculate the distribution
#     # of group circles. The circles are placed and distributed left
#     # to right around the center circle.
#     angle_list = []
#     # start_point needs to come down as numbers increase
#     start_point = 290 * ((29 - ((len(group)) - 2)) / 29)
#     # distribution increase needs to come down as numbers increase
#     distribution_start = 0
#     distribution_increase = 140
#     if len(group) > 2:  # no change in distribution increase if group of two
#         for i in range(len(group)):
#             distribution_increase = distribution_increase * 0.82
#     for i in range(len(group)):
#         angle = distribution_start + start_point
#         if angle > 360:
#             angle = angle - 360
#         angle_list.append(int(angle))
#         distribution_start += distribution_increase
#     assert isinstance(angle_list, (list,))
#
#
# @pytest.mark.skip(reason="Old. Pickle not used")
# def test_opening_a_pickle(master_pickle_file_path):
#     mickle = open_pickle_file(master_pickle_file_path)
#     assert str(mickle.master_data[0].quarter) == "Q1 20/21"
#
#
# @pytest.mark.skip(reason="old. Pickle not in use.")
# def test_master_in_a_pickle(full_test_masters_dict, project_info):
#     master = Master(full_test_masters_dict, project_info)
#     path_str = str("{0}/resources/test_master".format(os.path.join(os.getcwd())))
#     mickle = Pickle(master, path_str)
#     assert str(mickle.master.master_data[0].quarter) == "Q1 20/21"
#
#
# @pytest.mark.skip(reason="baselining not currently in use")
# def test_getting_baseline_data_from_masters(basic_masters_dicts, project_info):
#     master = JsonMaster(basic_masters_dicts, project_info)
#     master.get_baseline_data()
#     assert isinstance(master.bl_index, (dict,))
#     assert master.bl_index["ipdc_milestones"]["Sea of Tranquility"] == [0, 1]
#     assert master.bl_index["ipdc_costs"]["Apollo 11"] == [0, 1, 0, 2]
#     assert master.bl_index["ipdc_costs"]["Columbia"] == [0, 1, 0, 2]
