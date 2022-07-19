import argparse
from argparse import RawTextHelpFormatter
import sys
from typing import Dict, List
from pathlib import Path
import configparser

from docx.shared import Inches

from analysis_engine import __version__
from analysis_engine.core_data import get_core, open_json_file
from analysis_engine.dandelion import DandelionData, make_a_dandelion_auto
from analysis_engine.dashboards import narrative_dashboard, cdg_dashboard
from analysis_engine.dca import DcaData, dca_changes_into_word
from analysis_engine.render_utils import get_input_doc, put_matplotlib_fig_into_word
from analysis_engine.settings import report_config, set_default_args, return_koi_fn_keys
from analysis_engine.milestones import MilestoneData, milestone_chart, put_milestones_into_wb

# from analysis_engine.ar_data import get_gmpp_ar_data

# from analysis_engine.data import (
#     # get_master_data,
#     # Master,
#     # get_project_information,
#     VfMData,
#     # root_path,
#     # cdg_root_path,
#     vfm_into_excel,
#     MilestoneData,
#     put_milestones_into_wb,
#     run_p_reports,
#     RiskData,
#     risks_into_excel,
#     DcaData,
#     dca_changes_into_excel,
#     dca_changes_into_word,
#     ipdc_dashboard,
#     CostData,
#     cost_v_schedule_chart_into_wb,
#     # DandelionData,
#     # put_matplotlib_fig_into_word,
#     data_query_into_wb,
#     get_data_query_key_names,
#     # ProjectNameError,
#     # ProjectGroupError,
#     # ProjectStageError,
#     # milestone_chart,
#     cost_stackplot_graph,
#     # make_a_dandelion_auto,
#     build_speedials,
#     get_sp_data,
#     # DFT_GROUP,
#     # get_input_doc,
#     # InputError,
#     # JsonMaster,
#     # JsonData,
#     open_json_file,
#     cost_profile_into_wb_new,
#     cost_profile_graph_new,
#     # get_gmpp_data,
#     # place_gmpp_online_keys_into_dft_master_format,
#     portfolio_risks_into_excel,
# )

# from analysis_engine.ar_data import get_ar_data, ar_run_p_reports
from analysis_engine.gmpp_int import get_gmpp_data

from analysis_engine.error_msgs import (
    logger,
    ConfigurationError,
    ProjectNameError,
    InputError,
)


# As more and more meta data going into config could use a refactor to collect it all in
# one go?
from analysis_engine.speed_dials import build_speed_dials


def get_remove_income_totals(
    confi_path: Path,
) -> Dict:
    # Returns a list of dft groups
    try:
        config = configparser.ConfigParser()
        config.read(confi_path)
        dict = {
            "remove income from totals": config["COSTS"]["remove_income"],
        }
    except:
        logger.critical(
            "Configuration file issue. Please check remove_income list in the COST section"
        )
        sys.exit(1)

    return dict


def check_remove(op_args):  # subcommand arg
    if "remove" in op_args:
        from analysis_engine.data import CURRENT_LOG

        for p in op_args["remove"]:
            if p + " successfully removed from analysis." not in CURRENT_LOG:
                logger.warning(
                    p + " not recognised and therefore not removed from analysis."
                    ' Please make sure "remove" entry is correct.'
                )


def settings_switch(parse_args, report_type):
    """ "
    This function either runs the initiate function which saves core_data into a json file,
    or runs the run_analysis function which produces analytical analysis.
    """
    arguments = parse_args
    settings = report_config(report_type)
    if vars(arguments)["subparser_name"] == "initiate":
        initiate(settings)
    else:
        run_analysis(vars(arguments), settings)


def initiate(settings_dict):
    """
    This function does the 'initiate' command for all reports
    """
    logger.info("Initiating process to create master data for reporting.")
    get_core(settings_dict)
    logger.info(
        "The latest master and project information match. "
        "Master data file has been successfully created."
    )


def run_analysis(args, settings):
    programme = args["subparser_name"]
    op_args = {k: v for k, v in args.items() if v is not None}

    if ('dca', 'speed_dials', 'dashboards') == programme:
        op_args['quarter'] = 'standard'

    md = open_json_file(
        str(settings["root_path"]) + settings["master_path"],
        **op_args,
    )
    set_default_args(op_args, group=md["groups"], quarters=md["current_quarter"])
    combined_args = {**op_args, **settings}
    wb_save = False

    try:
        if programme == "dandelion":
            d_data = DandelionData(md, **combined_args)
            if op_args["chart"] != "save":
                make_a_dandelion_auto(d_data, **op_args)
            else:
                d_graph = make_a_dandelion_auto(d_data, **combined_args)
                doc_path = (
                    str(combined_args["root_path"]) + combined_args["word_landscape"]
                )
                doc = get_input_doc(doc_path)
                put_matplotlib_fig_into_word(doc, d_graph, width=Inches(8))
                doc_output_path = (
                    str(combined_args["root_path"]) + combined_args["word_save_path"]
                )
                doc.save(doc_output_path.format("dandelion"))

        if programme == "speed_dials":
            combined_args["rag_number"] = "5"
            # combined_args["quarter"] = "standard"
            sdmd = DcaData(md, **combined_args)
            sdmd.get_changes()
            sd_doc = get_input_doc(
                str(settings["root_path"]) + settings["word_landscape"]
            )
            build_speed_dials(sdmd, sd_doc)
            sd_doc.save(
                str(settings["root_path"])
                + settings["word_save_path"].format("speed_dials")
            )

        if programme == "dcas":
            # combined_args["quarter"] = "standard"
            sdmd = DcaData(md, **combined_args)
            sdmd.get_changes()
            changes_doc = dca_changes_into_word(
                sdmd, str(settings["root_path"]) + settings["word_portrait"]
            )
            changes_doc.save(
                str(settings["root_path"])
                + settings["word_save_path"].format("dca_changes")
            )

        if programme == "dashboards":
            narrative_d_master = get_input_doc(
                str(settings["root_path"]) + settings["narrative_dashboard"]
            )
            narrative_dashboard(md, narrative_d_master)  #
            narrative_d_master.save(
                str(settings["root_path"])
                + settings["excel_save_path"].format("narrative_dashboard_completed")
            )
            cdg_d_master = get_input_doc(
                str(settings["root_path"]) + settings["dashboard"]
            )
            cdg_dashboard(md, cdg_d_master)
            cdg_d_master.save(
                str(settings["root_path"])
                + settings["excel_save_path"].format("dashboard_completed")
            )

        if programme == "milestones":
            ms = MilestoneData(md, **combined_args)
            if (
                    # "type" in combined_args  # NOT IN USE.
                    "dates" in combined_args
                    or "koi" in combined_args
                    or "koi_fn" in combined_args
            ):
                return_koi_fn_keys(combined_args)
                ms.filter_chart_info(**combined_args)

            if op_args["chart"] != "save":
                ms_graph = milestone_chart(ms, **combined_args)
                doc = get_input_doc(
                    str(combined_args["root_path"]) + combined_args["word_landscape"]
                )
                put_matplotlib_fig_into_word(doc, ms_graph, width=Inches(8))
                doc.save(
                    str(combined_args["root_path"])
                    + combined_args["word_save_path"].format("milestones")
                )
            else:
                milestone_chart(ms, **combined_args)

            wb = put_milestones_into_wb(ms)
            wb_save = True

        if wb_save:
            if programme != "dashboards":
                wb.save(combined_args["root_path"] + "/output/{}.xlsx".format(programme))

    except (ProjectNameError, FileNotFoundError, InputError) as e:
        logger.critical(e)
        sys.exit(1)


# def ipdc_run_general(args):
#     # get portfolio reporting group information.
#     META = get_group_stage_data(
#         str(root_path) + "/core_data/ipdc_config.ini",
#     )
#     dft_group = META[0]
#     dft_stage = META[2]
#
#     programme = args["subparser_name"]
#     # wrap this into logging
#     try:
#         print("compiling ipdc " + programme + " analysis")
#     except TypeError:  # NoneType as no programme entered
#         print("Further command required. Use --help flag for guidance")
#         sys.exit(1)
#
#     m = Master(open_json_file(str(root_path / "core_data/json/master.json")))
#
#     try:
#         # print(args.items())
#         op_args = {
#             k: v for k, v in args.items() if v is not None
#         }  # removes None values
#         # print(op_args)
#         if "group" not in op_args:
#             if "stage" not in op_args:
#                 op_args["group"] = dft_group
#             if "stage" in op_args:
#                 if op_args["stage"] == []:
#                     op_args["stage"] = dft_stage
#         if "quarter" not in op_args:
#             if "baseline" not in op_args:
#                 op_args["quarter"] = ["standard"]
#
#         # projects to have income removed added
#         remove_income = get_remove_income_totals(
#             str(root_path) + "/core_data/ipdc_config.ini"
#         )
#         op_args["remove income from totals"] = remove_income[
#             "remove income from totals"
#         ]
#
#         # print(op_args)
#
#         if programme == "vfm":
#             c = VfMData(m, **op_args)  # c is class
#             wb = vfm_into_excel(c)
#
#         if programme == "risks":
#             c = RiskData(m, **op_args)
#             wb = risks_into_excel(c)
#
#         if programme == "portfolio_risks":
#             c = RiskData(m, **op_args)
#             wb = portfolio_risks_into_excel(c)
#
#         if programme == "dcas":
#             c = DcaData(m, **op_args)
#             wb = dca_changes_into_excel(c)
#
#         if programme == "costs":
#             if "baseline" in op_args:
#                 if op_args["baseline"] == ["current"]:
#                     op_args["quarter"] = [str(m.current_quarter)]
#                     c = CostData(m, **op_args)
#                     c.get_forecast_cost_profile()
#                     c.get_baseline_cost_profile()
#             else:
#                 c = CostData(m, **op_args)
#                 c.get_forecast_cost_profile()
#             wb = cost_profile_into_wb_new(c)
#             if "chart" not in op_args:
#                 op_args["chart"] = True
#                 cost_profile_graph_new(c, m, **op_args)
#             else:
#                 if op_args["chart"] == "save":
#                     op_args["chart"] = False
#                     cost_graph = cost_profile_graph_new(c, m, **op_args)
#                     doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
#                     put_matplotlib_fig_into_word(doc, cost_graph, size=7.5)
#                     doc.save(root_path / "output/costs_graph.docx")
#                 if op_args["chart"] == "show":
#                     op_args["chart"] = True
#                     cost_profile_graph_new(c, m, **op_args)
#
#         if programme == "costs_sp":
#             sp_data = get_sp_data(m, **op_args)
#
#             if "chart" not in op_args:
#                 op_args["chart"] = True
#                 cost_stackplot_graph(sp_data, m, **op_args)
#             else:
#                 if op_args["chart"] == "save":
#                     op_args["chart"] = False
#                     sp_graph = cost_stackplot_graph(sp_data, m, **op_args)
#                     doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
#                     put_matplotlib_fig_into_word(doc, sp_graph, size=7.5)
#                     doc.save(root_path / "output/stack_plot_graph.docx")
#                 if op_args["chart"] == "show":
#                     op_args["chart"] = True
#                     cost_stackplot_graph(sp_data, m, **op_args)
#
#         if programme == "speedial":
#             doc = get_input_doc(root_path / "input/summary_temp.docx")
#             land_doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
#             # if "conf_type" in op_args:
#             #     if op_args["conf_type"] == "sro_three":
#             #         op_args["rag_number"] = "3"
#             #         op_args["quarter"] = [str(m.current_quarter)]
#             #         data = DcaData(m, **op_args)
#             #         build_speedials(data, land_doc)
#             #         land_doc.save(root_path / "output/speed_dial_graph.docx")
#             #     else:  # refactor!!
#             op_args["rag_number"] = "3"
#             data = DcaData(m, **op_args)
#             data.get_changes()
#             doc = dca_changes_into_word(data, doc)
#             doc.save(root_path / "output/speed_dials_text.docx")
#             build_speedials(data, land_doc)
#             land_doc.save(root_path / "output/speed_dial_graph.docx")
#             # else:
#             #     op_args["rag_number"] = "5"
#             #     data = DcaData(m, **op_args)
#             #     data.get_changes()
#             #     doc = dca_changes_into_word(data, doc)
#             #     doc.save(root_path / "output/speed_dials_text.docx")
#             #     build_speedials(data, land_doc)
#             #     land_doc.save(root_path / "output/speed_dial_graph.docx")
#
#         if programme == "milestones":
#             ms = MilestoneData(m, **op_args)
#
#             if (
#                     "type" in op_args
#                     or "dates" in op_args
#                     or "koi" in op_args
#                     or "koi_fn" in op_args
#             ):
#                 op_args = return_koi_fn_keys(op_args)
#                 ms.filter_chart_info(**op_args)
#
#             if "chart" not in op_args:
#                 pass
#             else:
#                 if op_args["chart"] == "save":
#                     op_args["chart"] = False
#                     ms_graph = milestone_chart(ms, m, **op_args)
#                     doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
#                     put_matplotlib_fig_into_word(
#                         doc, ms_graph, size=8, transparent=False
#                     )
#                     doc.save(root_path / "output/milestones_chart.docx")
#                 if op_args["chart"] == "show":
#                     milestone_chart(ms, m, **op_args)
#
#             wb = put_milestones_into_wb(ms)
#
#         if programme == "dandelion":
#             if op_args["quarter"] == [
#                 "standard"
#             ]:  # converts "standard" default to current quarter
#                 op_args["quarter"] = [str(m.current_quarter)]
#             # op_args["order_by"] = "schedule"
#             d_data = DandelionData(m, **op_args)
#             if "chart" not in op_args:
#                 op_args["chart"] = True
#                 make_a_dandelion_auto(d_data, **op_args)
#             else:
#                 if op_args["chart"] == "save":
#                     op_args["chart"] = False
#                     d_graph = make_a_dandelion_auto(d_data, **op_args)
#                     doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
#                     put_matplotlib_fig_into_word(doc, d_graph, size=7)
#                     doc.save(root_path / "output/dandelion_graph.docx")
#                 if op_args["chart"] == "show":
#                     make_a_dandelion_auto(d_data, **op_args)
#
#         if programme == "dashboards":
#             # op_args["baseline"] = ["standard"]
#             dashboard_master = get_input_doc(root_path / "input/dashboards_master.xlsx")
#             wb = ipdc_dashboard(m, dashboard_master, op_args)
#             wb.save(root_path / "output/completed_ipdc_dashboard.xlsx")
#
#         if programme == "summaries":
#             op_args["quarter"] = [str(m.current_quarter)]
#             if "type" not in op_args:
#                 op_args["type"] = "short"
#             run_p_reports(m, **op_args)
#
#         # if programme == "top_250_summaries":
#         #     if op_args["group"] == dft_group:
#         #         op_args["group"] = ["HSRG", "RSS", "RIG", "RPE"]
#         #     top35_run_p_reports(m, **op_args)
#
#         if programme == "matrix":
#             costs = CostData(m, **op_args)
#             miles = MilestoneData(m, *op_args)
#             miles.calculate_schedule_changes()
#             wb = cost_v_schedule_chart_into_wb(miles, costs)
#             wb.save(root_path / "output/costs_schedule_matrix.xlsx")
#
#         if programme == "query":
#             if "koi" not in op_args and "koi_fn" not in op_args:
#                 logger.critical(
#                     "Please enter a key name(s) using either --koi or --koi_fn"
#                 )
#                 sys.exit(1)
#             op_args = return_koi_fn_keys(op_args)
#             wb = data_query_into_wb(m, **op_args)
#
#         if programme == "gmpp_data":
#             get_gmpp_data()
#
#         # if programme == "gmpp_ar":
#         #     ar_meta = get_gmpp_ar_data(
#         #         str(root_path) + "/core_data/ipdc_config.ini",
#         #     )
#         #     ar_data = get_ar_data(ar_meta["gmpp_ar_master"])
#         #     ar_run_p_reports(ar_data)
#
#         check_remove(op_args)
#
#         try:
#             if programme != "dashboards":
#                 wb.save(root_path / "output/{}.xlsx".format(programme))
#         except UnboundLocalError:
#             pass
#
#         print(programme + " analysis has been compiled. Enjoy!")
#
#     except (ProjectNameError, FileNotFoundError, InputError) as e:
#         logger.critical(e)
#         sys.exit(1)
#
#     # TODO optional_args produces a list of strings, each of which are to be in the output file name path.
#     # optional_args = get_args_for_file(args)
#     # wb.save(root_path / "output/{}_{}.xlsx".format(programme, optional_args))
#     # print(programme + " analysis has been compiled. Enjoy!")
#
# def cdg_run_general(args):

#

#
#
#         if programme == "dashboards":
#             dashboard_master = get_input_doc(
#                 cdg_root_path / "input/dashboard_master.xlsx"
#             )
#             narrative_d_master = get_input_doc(
#                 cdg_root_path / "input/narrative_dashboard_master.xlsx"
#             )
#             wb = cdg_narrative_dashboard(m, narrative_d_master)
#             wb.save(
#                 cdg_root_path
#                 / "output/{}.xlsx".format("cdg_narrative_dashboard_completed")
#             )
#             wb = cdg_dashboard(m, dashboard_master)
#             wb.save(cdg_root_path / "output/{}.xlsx".format("cdg_dashboard_completed"))
#
#         check_remove(op_args)
#
#         try:
#             if programme != "dashboards":
#                 wb.save(cdg_root_path / "output/{}.xlsx".format(programme))
#         except UnboundLocalError:
#             pass
#
#         print(programme + " analysis has been compiled. Enjoy!")
#
#     except (ProjectNameError, FileNotFoundError, InputError) as e:
#         logger.critical(e)
#         sys.exit(1)
#


def run_parsers():
    report_type = sys.argv[1]

    parser = argparse.ArgumentParser(
        description=f"runs all analysis for {report_type} reporting"
    )
    subparsers = parser.add_subparsers(dest="subparser_name")
    subparsers.metavar = "                      "
    parser_initiate = subparsers.add_parser(
        "initiate", help="creates a master data file"
    )
    parser_milestones = subparsers.add_parser(
        "milestones",
        help="milestone schedule graphs and data.",
    )
    dandelion_description = (
        "Creates the 'dandelion' graph. See below optional arguments for changing the "
        "dandelion that is compiled. The command analysis dandelion returns the default "
        'dandelion graph. The user must specify --chart "save" to save the chart, otherwise '
        "only a temporary chart will be generated."
    )
    parser_dandelion = subparsers.add_parser(
        "dandelion",
        help="Dandelion graph.",
        description=dandelion_description,
    )
    dashboard_description = (
        "Creates dashboards. There are no optional arguments for this command.\n\n"
        "A blank master dashboard titled dashboards_master.xlsx must be in input file.\n\n"
        "A completed dashboard named completed_dashboard.xlsx will be placed into\n"
        "the output file."
    )
    subparsers.add_parser(
        "dashboards",
        help="dashboard",
        description=dashboard_description,
        formatter_class=RawTextHelpFormatter,
    )  # no associated op args.

    parser_speed_dial = subparsers.add_parser("speed_dials", help="speed dial analysis")
    parser_dca = subparsers.add_parser("dcas", help="dca analysis")

    for sub in [
        parser_milestones,
        parser_dandelion,
    ]:
        sub.add_argument(
            "--chart",
            type=str,
            metavar="",
            action="store",
            choices=["show", "save"],
            help="options for building and saving graph output. Commands are 'show' or 'save' ",
        )

    # quarter
    for sub in [
        parser_speed_dial,
        parser_dandelion,
        parser_dca,
        parser_milestones,
    ]:
        sub.add_argument(
            "--quarter",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            help="Returns analysis for one or combination of specified quarters. "
            'User must use correct format e.g "Q3 19/20"',
        )

    # stage
    for sub in [
        parser_dca,
        parser_speed_dial,
        # parser_vfm,
        # parser_risks,
        # parser_port_risks,
        parser_dandelion,  # ipdc also has pipeline option. Not tested yet.
        # parser_costs,
        # parser_costs_sp,
        # parser_data_query,
        parser_milestones,
        # parser_data_query,
    ]:
        sub.add_argument(
            "--stage",
            type=str,
            metavar="",
            action="store",
            nargs="*",
            choices=["FBC", "OBC", "SOBC", "pre-SOBC", "pipeline"],
            help="Returns analysis for only those projects at the specified planning stage(s). By default "
            "the --stage argument will return the list of bc_stages specified in the config file."
            'Or user can enter one or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
        )
    # group
    for sub in [
        parser_dca,
        # parser_vfm,
        # parser_risks,
        # parser_port_risks,
        parser_speed_dial,
        parser_dandelion,
        # parser_costs,
        parser_milestones,
        # parser_summaries,
        # parser_costs_sp,
        # parser_data_query,
    ]:
        sub.add_argument(
            "--group",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            help="Returns analysis for specified project(s), only. User must enter one or a combination of "
            'DfT Group names; "HSRG", "RSS", "RIG", "AMIS","RPE", or the project(s) acronym or full name.',
        )

    parser_dandelion.add_argument(
        "--angles",
        type=int,
        metavar="",
        action="store",
        nargs="+",
        # choices=['sro', 'finance', 'benefits', 'schedule', 'resource'],
        help="Use can manually enter angles for group bubbles",
    )

    parser_dandelion.add_argument(
        "--type",
        type=str,
        metavar="",
        action="store",
        choices=["benefits", "income"],
        help="Provide the type of value to include in dandelion. Options are"
        ' "benefits" or "income".',
    )

    parser_milestones.add_argument(
        "--dates",
        type=str,
        metavar="",
        action="store",
        nargs=2,
        help="dates for analysis. Must provide start date and then end date in format e.g. '1/1/2021' '1/1/2022'.",
    )

    parser_milestones.add_argument(
        "--blue_line",
        type=str,
        metavar="",
        action="store",
        help="Insert blue line into chart to represent a date. "
             'Options are "Today" "CDG" or a date in correct format e.g. "1/1/2021".',
    )

    cli_args = parser.parse_args(sys.argv[2:])
    settings_switch(cli_args, report_type)


class main:
    def __init__(self):
        ae_description = (
            "Welcome to the DfT Major Projects Portfolio Office analysis engine.\n\n"
            "To operate use subcommands outlined below. To navigate each subcommand\n"
            "option use the --help flag which will provide instructions on which optional\n"
            "arguments can be used with each subcommand. e.g. analysis dandelion --help."
        )
        parser = argparse.ArgumentParser(
            description=ae_description, formatter_class=RawTextHelpFormatter
        )

        parser.add_argument("--version", action="version", version=__version__)

        parser.add_argument(
            "command",
            help="Initial command to specify whether ipdc, top2_50, cdg analysis if required",
        )

        args = parser.parse_args(sys.argv[1:2])
        if vars(args)["command"] not in ["ipdc", "top_250", "cdg"]:
            print("Unrecognised command. Options are ipdc, top250 or cdg")
            exit(1)

        # use dispatch pattern to invoke method with same name
        getattr(self, args.command)()

    def ipdc(self):
        parser = argparse.ArgumentParser(
            description="runs all analysis for ipdc reporting"
        )
        subparsers = parser.add_subparsers(dest="subparser_name")
        subparsers.metavar = "                      "
        # parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
        parser_initiate = subparsers.add_parser(
            "initiate", help="creates a master data file"
        )
        dashboard_description = (
            "Creates IPDC dashboards. There are no optional arguments for this command.\n\n"
            "A blank master dashboard titled dashboards_master.xlsx must be in input file.\n\n"
            "A completed dashboard title completed_ipdc_dashboard.xlsx will be placed into\n"
            "the output file."
        )
        parser_dashboard = subparsers.add_parser(
            "dashboards",
            help="IPDC dashboards",
            description=dashboard_description,
            formatter_class=RawTextHelpFormatter,
        )
        dandelion_description = (
            "Creates the IPDC 'dandelion' graph. See below optional arguments for changing the "
            "dandelion that is compiled. The command analysis dandelion returns the default "
            'dandelion graph. The user must specify --chart "save" to save the chart, otherwise '
            "only a temporary matplotlib chart will be generated."
        )
        parser_dandelion = subparsers.add_parser(
            "dandelion",
            help="Dandelion graph.",
            description=dandelion_description,
            # formatter_class=RawTextHelpFormatter,  # can't use as effects how optional arguments are shown.
        )

        costs_description = (
            "Creates a cost profile graph. See below optional arguments. The user "
            'must specify --chart "save" to save the chart, otherwise '
            "only a temporary matplotlib chart will be generated."
        )

        parser_costs = subparsers.add_parser(
            "costs",
            help="cost trend profile graph and data.",
            description=costs_description,
        )

        costs_sp_description = (
            "Creates a cost stack plot profile graph. See below optional arguments. The user "
            'must specify --chart "save" to save the chart, otherwise '
            "only a temporary matplotlib chart will be generated."
        )

        parser_costs_sp = subparsers.add_parser(
            "costs_sp",
            help="cost stack plot graph and data.",
            description=costs_sp_description,
        )

        parser_milestones = subparsers.add_parser(
            "milestones",
            help="milestone schedule graphs and data.",
        )
        parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
        parser_summaries = subparsers.add_parser("summaries", help="summary reports")
        parser_risks = subparsers.add_parser("risks", help="project risk analysis")
        parser_port_risks = subparsers.add_parser(
            "portfolio_risks", help="portfolio risk analysis"
        )
        parser_dca = subparsers.add_parser("dcas", help="dca analysis")
        parser_speedial = subparsers.add_parser("speedial", help="speed dial analysis")
        parser_matrix = subparsers.add_parser(
            "matrix", help="cost v schedule chart. In development not working."
        )
        parser_data_query = subparsers.add_parser(
            "query", help="return data from core data"
        )
        parser_gmpp_data = subparsers.add_parser(
            "gmpp_data", help="converts gmpp online data into the dft master format"
        )
        parser_gmpp_ar = subparsers.add_parser(
            "gmpp_ar", help="compiled summaries for the IPA GMPP annual report"
        )

        # Arguments
        # stage
        for sub in [
            parser_dca,
            parser_vfm,
            parser_risks,
            parser_port_risks,
            parser_speedial,
            # parser_dandelion,
            parser_costs,
            parser_costs_sp,
            # parser_data_query,
            parser_milestones,
            parser_data_query,
        ]:
            sub.add_argument(
                "--stage",
                type=str,
                metavar="",
                action="store",
                nargs="*",
                choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
                help="Returns analysis for only those projects at the specified planning stage(s). By default "
                "the --stage argument will return the list of bc_stages specified in the config file."
                'Or user can enter one or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
            )
        # stage dandelion only
        parser_dandelion.add_argument(
            "--stage",
            type=str,
            metavar="",
            action="store",
            nargs="*",
            choices=["FBC", "OBC", "SOBC", "pre-SOBC", "pipeline"],
            help="Returns analysis for only those projects at the specified planning stage(s). By default "
            "the --stage argument will return the list of bc_stages specified in the config file."
            'Or user can enter one or combination of "FBC", "OBC", "SOBC", "pre-SOBC".For dandelion "pipeline" is also available',
        )

        # group
        for sub in [
            parser_dca,
            parser_vfm,
            parser_risks,
            parser_port_risks,
            parser_speedial,
            parser_dandelion,
            parser_costs,
            parser_milestones,
            parser_summaries,
            parser_costs_sp,
            parser_data_query,
        ]:
            sub.add_argument(
                "--group",
                type=str,
                metavar="",
                action="store",
                nargs="+",
                help="Returns analysis for specified project(s), only. User must enter one or a combination of "
                'DfT Group names; "HSRG", "RSS", "RIG", "AMIS","RPE", or the project(s) acronym or full name.',
            )
        # remove
        for sub in [
            parser_dca,
            parser_vfm,
            parser_risks,
            parser_port_risks,
            parser_speedial,
            parser_dandelion,
            parser_costs,
            parser_costs_sp,
            parser_data_query,
            parser_milestones,
        ]:
            sub.add_argument(
                "--remove",
                type=str,
                metavar="",
                action="store",
                nargs="+",
                help="Removes specified projects from analysis. User must enter one or a combination of either"
                " a recognised DfT Group name, a recognised planning stage or the project(s) acronym or full"
                " name.",
            )
        # quarter
        for sub in [
            parser_dca,
            parser_vfm,
            parser_risks,
            parser_port_risks,
            parser_speedial,
            parser_dandelion,
            parser_costs,
            parser_costs_sp,
            parser_data_query,
            parser_milestones,
        ]:
            sub.add_argument(
                "--quarter",
                type=str,
                metavar="",
                action="store",
                nargs="+",
                help="Returns analysis for one or combination of specified quarters. "
                'User must use correct format e.g "Q3 19/20"',
            )

        parser_costs.add_argument(
            "--baseline",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            choices=["current"],
            help="baseline option for costs refactored in Q1 21/22. Choose 'current' to return project "
            "reported bls as well as latest forecast profile",
        )

        parser_milestones.add_argument(
            "--type",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            choices=["Approval", "Assurance", "Delivery"],
            help="Returns analysis for specified type of milestones.",
        )

        parser_speedial.add_argument(
            "--conf_type",
            type=str,
            metavar="",
            action="store",
            choices=["sro", "finance", "benefits", "schedule", "resource"],
            help="Returns analysis for specified confidence types. options are"
            "'sro', 'finance', 'benefits', 'schedule', 'resource'."
            " As of Q2 20/21 it only provides a three rag dial.",
        )

        for sub in [parser_milestones, parser_data_query]:
            sub.add_argument(
                "--koi",
                type=str,
                action="store",
                nargs="+",
                help="Returns the specified keys of interest (KOI).",
            )

        for sub in [parser_milestones, parser_data_query]:
            sub.add_argument(
                "--koi_fn",
                type=str,
                action="store",
                help="provide name of csv file contain key names",
            )

        parser_milestones.add_argument(
            "--dates",
            type=str,
            metavar="",
            action="store",
            nargs=2,
            help="dates for analysis. Must provide start date and then end date in format e.g. '1/1/2021' '1/1/2022'.",
        )

        parser_dandelion.add_argument(
            "--type",
            type=str,
            metavar="",
            action="store",
            choices=[
                "spent",
                "remaining",
                "benefits",
                "ps resource",
                "contract resource",
                "total resource",
                "funded resource",
            ],
            help="Provide the type of value to include in dandelion. Options are"
            ' "spent", "remaining", "benefits", "ps resource", "contract resource", "total resource", "funded resource".',
        )

        parser_dandelion.add_argument(
            "--order_by",
            type=str,
            metavar="",
            action="store",
            choices=["schedule"],
            help="Specify how project circles should be ordered: 'schedule' only current"
            " option.",
        )

        parser_summaries.add_argument(
            "--type",
            type=str,
            metavar="",
            action="store",
            choices=["long", "short"],
            help="Specify which form of report is required. Options are 'short' or 'long'",
        )

        parser_costs_sp.add_argument(
            "--type",
            type=str,
            metavar="",
            action="store",
            choices=["cat"],
            help="Provide the type of value to include in dandelion. Options are"
            ' "cat".',
        )

        parser_dandelion.add_argument(
            "--angles",
            type=int,
            metavar="",
            action="store",
            nargs="+",
            # choices=['sro', 'finance', 'benefits', 'schedule', 'resource'],
            help="Use can manually enter angles for group bubbles",
        )

        parser_dandelion.add_argument(
            "--confidence",
            type=str,
            metavar="",
            action="store",
            choices=["sro", "finance", "benefits", "schedule", "resource"],
            help="specify the confidence type to displayed for each project.",
        )

        parser_dandelion.add_argument(
            "--pc",
            type=str,
            metavar="",
            action="store",
            choices=["G", "A/G", "A", "A/R", "R"],
            help="specify the colour for the overall portfolio circle",
        )

        parser_dandelion.add_argument(
            "--circle_colour",
            type=str,
            metavar="",
            action="store",
            choices=["No", "Yes"],
            help="specify whether to colour circles with DCA rating colours",
        )

        parser_dandelion.add_argument(
            "--circle_edge",
            type=str,
            metavar="",
            action="store",
            choices=["forward_look", "ipa"],
            help="specify whether to colour circle edge with SRO forward look rating. "
            "Options are 'forward_look' or 'ipa'.",
        )

        # chart
        for sub in [parser_dandelion, parser_costs, parser_costs_sp, parser_milestones]:
            sub.add_argument(
                "--chart",
                type=str,
                metavar="",
                action="store",
                choices=["show", "save"],
                help="options for building and saving graph output. Commands are 'show' or 'save' ",
            )

        # title
        for sub in [parser_costs, parser_milestones, parser_costs_sp]:
            sub.add_argument(
                "--title",
                type=str,
                metavar="",
                action="store",
                help="provide a title for chart. Optional",
            )

        parser_milestones.add_argument(
            "--blue_line",
            type=str,
            metavar="",
            action="store",
            help="Insert blue line into chart to represent a date. "
            'Options are "today" "config_date" or a date in correct format e.g. "1/1/2021".',
        )

        cli_args = parser.parse_args(sys.argv[2:])
        settings_switch(cli_args, "ipdc")

    def cdg(self):
        run_parsers()


if __name__ == "__main__":
    main()
