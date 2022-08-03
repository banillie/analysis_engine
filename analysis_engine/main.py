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
from analysis_engine.dashboards import narrative_dashboard, cdg_dashboard, ipdc_dashboard
from analysis_engine.dca import DcaData, dca_changes_into_word
from analysis_engine.gmpp_int import GmppOnlineCosts
from analysis_engine.merge import Merge
from analysis_engine.query import data_query_into_wb
from analysis_engine.render_utils import get_input_doc, put_matplotlib_fig_into_word
from analysis_engine.risks import RiskData, portfolio_risks_into_excel, risks_into_excel
from analysis_engine.settings import (
    report_config,
    set_default_args,
    return_koi_fn_keys,
    get_integration_data,
    get_masters_to_merge,
)
from analysis_engine.milestones import (
    MilestoneData,
    milestone_chart,
    put_milestones_into_wb,
)
from analysis_engine.speed_dials import build_speed_dials
from analysis_engine.error_msgs import no_query_keys

from analysis_engine.error_msgs import (
    logger,
    ConfigurationError,
    ProjectNameError,
    InputError,
)


class CliOpArgs:
    def __init__(self, args, settings):
        self.args = args
        self.settings = settings
        self.combined_args = {}
        self.md = {}
        self.wb_save = False
        self.programme = ""
        self.cli_op_args()

    def cli_op_args(self):
        self.programme = self.args["subparser_name"]
        op_args = {k: v for k, v in self.args.items() if v is not None}

        # these programs have the latest two quarters as default.
        # other program defaults are setting very get_iter_list()
        if self.programme in ["dcas", "speed_dials"]:
            if "quarter" not in list(op_args.keys()):
                op_args["quarter"] = "standard"

        if self.programme == "milestones":
            if "chart" not in (op_args.keys()):
                op_args["chart"] = None

        if self.programme == "dashboards":
            op_args["quarter"] = "four"

        try:
            if self.programme == "query":
                if "koi" not in op_args and "koi_fn" not in op_args:
                    no_query_keys()
        except InputError as e:
            logger.critical(e)
            sys.exit(1)

        md = open_json_file(
            str(self.settings["root_path"]) + self.settings["master_path"],
            **op_args,
        )
        set_default_args(
            op_args,
            group=md["groups"],
            quarters=md["current_quarter"],
            stage=md["stages"],
        )
        combined_args = {**op_args, **self.settings}

        if combined_args["report"] == "ipdc":
            combined_args["circle_edge"] = "forward_look"  # for dandelion

        if self.programme == "gmpp_data":
            get_integration_data(combined_args)

        self.combined_args = combined_args
        self.md = md
        self.wb_save = False


def settings_switch(parse_args, report_type):
    """
    This function either runs the initiate function which saves core_data into a json file,
    or runs the run_analysis function which produces analytical analysis.
    """
    args = vars(parse_args)
    settings = report_config(report_type)
    if args["subparser_name"] == "initiate":
        initiate(settings)
    else:
        run_analysis(args, settings)


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
    cli = CliOpArgs(args, settings)
    try:
        if cli.programme == "dandelion":

            if cli.combined_args["report"] == "ipdc":
                cli.combined_args["abbreviations"] = True

            d_data = DandelionData(cli.md, **cli.combined_args)
            if cli.combined_args["chart"] != "save":
                make_a_dandelion_auto(d_data, **cli.combined_args)
            else:
                d_graph = make_a_dandelion_auto(d_data, **cli.combined_args)
                doc_path = (
                    str(cli.combined_args["root_path"])
                    + cli.combined_args["word_landscape"]
                )
                doc = get_input_doc(doc_path)
                put_matplotlib_fig_into_word(doc, d_graph, width=Inches(8))
                doc_output_path = (
                    str(cli.combined_args["root_path"])
                    + cli.combined_args["word_save_path"]
                )
                doc.save(doc_output_path.format("dandelion"))

        if cli.programme == "speed_dials":

            if cli.combined_args["report"] == "ipdc":
                cli.combined_args["rag_number"] = "3"
            if cli.combined_args["report"] == "cdg":
                cli.combined_args["rag_number"] = "5"

            sdmd = DcaData(cli.md, **cli.combined_args)
            sdmd.get_changes()
            sd_doc = get_input_doc(
                str(cli.combined_args["root_path"])
                + cli.combined_args["word_landscape"]
            )
            build_speed_dials(sdmd, sd_doc)
            sd_doc.save(
                str(cli.combined_args["root_path"])
                + cli.combined_args["word_save_path"].format("speed_dials")
            )

        if cli.programme == "dcas":
            sdmd = DcaData(cli.md, **cli.combined_args)
            sdmd.get_changes()
            changes_doc = dca_changes_into_word(
                sdmd,
                str(cli.combined_args["root_path"])
                + cli.combined_args["word_portrait"],
            )
            changes_doc.save(
                str(cli.combined_args["root_path"])
                + cli.combined_args["word_save_path"].format("dca_changes")
            )

        if cli.programme == "dashboards":
            if cli.combined_args['report'] == 'cdg':
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
            if cli.combined_args['report'] == 'ipdc':
                ipdc_d_master = get_input_doc(
                    str(cli.combined_args["root_path"]) + cli.combined_args["dashboard"]
                )
                ipdc_dashboard(cli.md, ipdc_d_master, **cli.combined_args)
                ipdc_d_master.save(
                    str(cli.combined_args["root_path"])
                    + cli.combined_args["excel_save_path"].format("dashboards_completed")
                )

        if cli.programme == "milestones":
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
                    str(cli.combined_args["root_path"])
                    + cli.combined_args["word_landscape"]
                )
                put_matplotlib_fig_into_word(doc, ms_graph, width=Inches(8))
                doc.save(
                    str(cli.combined_args["root_path"])
                    + cli.combined_args["word_save_path"].format("milestones_chart")
                )
            if cli.combined_args["chart"] == "show":
                milestone_chart(ms, **cli.combined_args)

            wb = put_milestones_into_wb(ms)
            wb.save(
                cli.combined_args["root_path"] + "/output/{}.xlsx".format(cli.programme)
            )

        if cli.programme == "query":
            op_args = return_koi_fn_keys(cli.combined_args)
            wb = data_query_into_wb(cli.md, **op_args)
            wb.save(
                str(cli.settings["root_path"])
                + settings["excel_save_path"].format("query")
            )

        if cli.programme == "gmpp_data":
            md = GmppOnlineCosts(**cli.combined_args)
            md.place_into_dft_master_format()

        if cli.programme == "merge_masters":
            get_masters_to_merge(cli.combined_args)
            Merge(**cli.combined_args)

        if cli.programme == "portfolio_risks":
            rd = RiskData(cli.md, **cli.combined_args)
            wb = portfolio_risks_into_excel(rd)
            wb.save(
                str(cli.settings["root_path"])
                + cli.settings["excel_save_path"].format("portfolio_risks")
            )

        if cli.programme == 'project_risks':
            rd = RiskData(cli.md, **cli.combined_args)
            wb = risks_into_excel(rd)
            wb.save(
                str(cli.settings["root_path"])
                + cli.settings["excel_save_path"].format("project_risks")
            )

        # ProjectNameError below captures any naming errors both this is hit.
        if "remove" in cli.combined_args.keys():
            for x in cli.combined_args["remove"]:
                logger.info(f"{x} has been removed from analysis")

    except (ProjectNameError, FileNotFoundError, InputError) as e:
        logger.critical(e)
        sys.exit(1)


def run_parsers():
    report_type = sys.argv[1]

    parser = argparse.ArgumentParser(
        description=f"Welcome to analysis engine for {report_type} reporting. All optional commands are below."
    )
    subparsers = parser.add_subparsers(dest="subparser_name")
    subparsers.metavar = "                      "

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
        "Creates dashboards. A blank master dashboard titled dashboards_master.xlsx must be in input folder. "
        "A completed dashboard named completed_dashboard.xlsx will be placed into "
        "the output file."
    )
    subparsers.add_parser(
        "dashboards",
        help="Dashboard compilation.",
        description=dashboard_description,
        formatter_class=RawTextHelpFormatter,
    )

    parser_dca = subparsers.add_parser(
        "dcas",
        help="DCA comparisons between quarters",
        description="Generates a print out on DCA changes between quarters. All DCAs are placed "
                    "into the same file and placed in output folder. It has a maximum of two quarters "
                    "for the --quarters argument.",
    )

    parser_gmpp_online = subparsers.add_parser(
        "gmpp_data",
        help="Compiles GMPP online data (as provided by IPA) into master file format.",
        description="This program converts gmpp online data into the dft master file "
                    "format. It requires three files. A file containing the gmpp online data - as provided by "
                    "the IPA, a GMPP_INTEGRATION_KEY_MAP, and a GMPP_INTEGRATION_PROJECT_MAP. These files must "
                    "be referenced in the config file in the GMPP INTEGRATION section and saved in the input folder. "
                    "There are no optional arguments for this program other than --help.",
    )

    parser_initiate = subparsers.add_parser(
        "initiate", help="IMPORTANT. This command must be run everytime new data is saved in core_data."
    )

    parser_merge_masters = subparsers.add_parser(
        "merge_masters",
        help="This program takes separate master data files and places them into one master file",
        description="This program takes separate master data files and places them into one master file. The user "
                    "can set the masters that it would like to merge via the config file in the MERGE / masters_list "
                    "section. The masters must be saved in the input folder. There are no optional arguments for this "
                    "program other than --help.",
    )

    parser_milestones = subparsers.add_parser(
        "milestones",
        help="Compiles milestone schedule graphs and data.",
        description="Generates raw data outputs as well as visuals for milestone data. The default is simply "
        "the return of an excel file with milestone data. Use the --chart options to produce"
        "visual outputs ",
    )

    parser_port_risks = subparsers.add_parser(
        "portfolio_risks",
        help="Compiles portfolio risk analysis.",
        description="This program extracts all portfolio risk data used in the IPDC portfolio report."
    )
    parser_risks = subparsers.add_parser(
        "project_risks",
        help="Compiles project risk analysis",
        description='This program extracts all project level risk data contained in the master data set.'
    )

    parser_data_query = subparsers.add_parser(
        "query", help="Returns data from master data set as specified by the user."
    )

    parser_speed_dial = subparsers.add_parser(
        "speed_dials",
        help="Produces the speed dial analysis and visuals used in the portfolio management report.",
        description="Creates the speed dial visual outputs. All confidence speed dials are created "
        "at the same time and saved into the output folder. It has a maximum of two quarters "
        "for the --quarters argument.",
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

    parser_milestones.add_argument(
        "--blue_line",
        type=str,
        metavar="",
        action="store",
        choices=["config_date", "today"],
        help="User can insert blue line into chart to represent a date. "
             "Options are 'config_date' or 'today'. The config_data is set in the config file in the "
             "Globals / milestones_blue_line_date value.",
    )

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
            help="Creates a graphical output. The user can choose to either temporarily 'show' the chart "
            "or 'save' it into a word document which is saved into the output folder",
        )

    parser_milestones.add_argument(
        "--dates",
        type=str,
        metavar="",
        action="store",
        nargs=2,
        help="dates for analysis. Must provide start date and then end date in format e.g. '1/1/2021' '1/1/2022'.",
    )

    # group
    for sub in [
        parser_dca,
        parser_risks,
        parser_port_risks,
        parser_speed_dial,
        parser_dandelion,
        parser_milestones,
        parser_data_query,
    ]:
        sub.add_argument(
            "--group",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            help="Returns analysis for specified project(s), only. User must enter one or a combination of "
                 "DfT Group names. Group names must match those in the config document. For the dandelion command the user "
                 "has an added group option of 'pipeline'.",
        )

    for sub in [parser_milestones, parser_data_query]:
        sub.add_argument(
            "--koi",
            type=str,
            action="store",
            nargs="+",
            help="Key of interest (koi). The user can specify data keys that are of specific interest for "
            "analysis. The user can specify keys as in the master document or they can specify "
            "milestone names being reported by projects.",
        )

    for sub in [parser_milestones, parser_data_query]:
        sub.add_argument(
            "--koi_fn",
            type=str,
            action="store",
            help="Key of interest file name (koi_fn). As per --koi. But in this instance the user can specify "
            "keys via a csv document saved in the input folder. Keys need to be in column A, and A1 should be "
            "titled key_name. ",
        )

    parser_dandelion.add_argument(
        "--order_by",
        type=str,
        metavar="",
        action="store",
        choices=["schedule"],
        help="User can change the order in which circles are placed. The only choice for "
             "this argument currently is 'schedule' ",
    )

    # quarter
    for sub in [
        parser_speed_dial,
        parser_dandelion,
        parser_dca,
        parser_milestones,
        parser_data_query,
        parser_port_risks,
        parser_risks,
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

    # remove
    for sub in [
        parser_dca,
        parser_risks,
        parser_port_risks,
        parser_speed_dial,
        parser_dandelion,
        parser_data_query,
        parser_milestones,
    ]:
        sub.add_argument(
            "--remove",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            help="The User can remove project(s) from analysis is necessary. The projects full name "
            "or acronym can be used.",
        )

    # stage
    for sub in [
        parser_dca,
        parser_speed_dial,
        parser_risks,
        parser_port_risks,
        parser_dandelion,
        parser_data_query,
        parser_milestones,
    ]:
        sub.add_argument(
            "--stage",
            type=str,
            metavar="",
            action="store",
            nargs="*",
            choices=["FBC", "OBC", "SOBC", "pre-SOBC", "pipeline"],
            help="Returns analysis for those projects at the specified planning stage(s). By default "
            "the --stage argument will return the list of business case stages specified in the config file. "
            "Or user can enter one or combination of business cases (which must match the those specified in "
            "the config file). The dandelion the dandelion command the user has the added option of 'pipeline'",
        )

    parser_milestones.add_argument(
        "--title",
        type=str,
        metavar="",
        action="store",
        help="The user can specify and title for the chart output. Please enter as text, for "
             "example 'This the title'.",
    )

    parser_dandelion.add_argument(
        "--type",
        type=str,
        metavar="",
        action="store",
        choices=[
            "spent_costs",
            "remaining_costs",
            "income",
            "ps_resource",
            "contractor_resource",
            "total_resource",
            "funded_resource",
        ],
        help="The user can specify the type of value for project bubble sizes in the dandelion. Options are "
        "'remaining_costs', 'spent_costs', 'income', 'ps_resource', 'contractor_resource', "
        "'total_resource' or 'funded_resource'.",
    )

    cli_args = parser.parse_args(sys.argv[2:])
    settings_switch(cli_args, report_type)


class main:
    def __init__(self):
        ae_description = (
            "Welcome to the DfT Major Projects Portfolio Office analysis engine.\n  \n"
            "To operate use subcommands outlined below. To navigate each subcommand "
            "option use the --help flag which will provide instructions on which optional "
            "arguments can be used with each subcommand. e.g. analysis dandelion --help."
        )
        parser = argparse.ArgumentParser(
            description=ae_description, formatter_class=RawTextHelpFormatter
        )

        parser.add_argument("--version", action="version", version=__version__)

        parser.add_argument(
            "command",
            help="Initial command to specify the reporting type required. Options are 'ipdc' or 'cdg'.",
        )

        args = parser.parse_args(sys.argv[1:2])
        if vars(args)["command"] not in ["ipdc", "cdg"]:
            print("Unrecognised command. Options are ipdc or cdg")
            exit(1)

        # use dispatch pattern to invoke method with same name
        getattr(self, args.command)()

    def ipdc(self):
        run_parsers()
        # parser = argparse.ArgumentParser(
        #     description="runs all analysis for ipdc reporting"
        # )
        # subparsers = parser.add_subparsers(dest="subparser_name")
        # subparsers.metavar = "                      "
        # # parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
        # parser_initiate = subparsers.add_parser(
        #     "initiate", help="creates a master data file"
        # )
        # dashboard_description = (
        #     "Creates IPDC dashboards. There are no optional arguments for this command.\n\n"
        #     "A blank master dashboard titled dashboards_master.xlsx must be in input file.\n\n"
        #     "A completed dashboard title completed_ipdc_dashboard.xlsx will be placed into\n"
        #     "the output file."
        # )
        # parser_dashboard = subparsers.add_parser(
        #     "dashboards",
        #     help="IPDC dashboards",
        #     description=dashboard_description,
        #     formatter_class=RawTextHelpFormatter,
        # )
        # dandelion_description = (
        #     "Creates the IPDC 'dandelion' graph. See below optional arguments for changing the "
        #     "dandelion that is compiled. The command analysis dandelion returns the default "
        #     'dandelion graph. The user must specify --chart "save" to save the chart, otherwise '
        #     "only a temporary matplotlib chart will be generated."
        # )
        # parser_dandelion = subparsers.add_parser(
        #     "dandelion",
        #     help="Dandelion graph.",
        #     description=dandelion_description,
        #     # formatter_class=RawTextHelpFormatter,  # can't use as effects how optional arguments are shown.
        # )
        #
        # costs_description = (
        #     "Creates a cost profile graph. See below optional arguments. The user "
        #     'must specify --chart "save" to save the chart, otherwise '
        #     "only a temporary matplotlib chart will be generated."
        # )
        #
        # parser_costs = subparsers.add_parser(
        #     "costs",
        #     help="cost trend profile graph and data.",
        #     description=costs_description,
        # )
        #
        # costs_sp_description = (
        #     "Creates a cost stack plot profile graph. See below optional arguments. The user "
        #     'must specify --chart "save" to save the chart, otherwise '
        #     "only a temporary matplotlib chart will be generated."
        # )
        #
        # parser_costs_sp = subparsers.add_parser(
        #     "costs_sp",
        #     help="cost stack plot graph and data.",
        #     description=costs_sp_description,
        # )
        #
        # parser_milestones = subparsers.add_parser(
        #     "milestones",
        #     help="milestone schedule graphs and data.",
        # )
        # parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
        # parser_summaries = subparsers.add_parser("summaries", help="summary reports")
        # parser_risks = subparsers.add_parser("risks", help="project risk analysis")
        # parser_port_risks = subparsers.add_parser(
        #     "portfolio_risks", help="portfolio risk analysis"
        # )
        # parser_dca = subparsers.add_parser("dcas", help="dca analysis")
        # parser_speedial = subparsers.add_parser("speedial", help="speed dial analysis")
        # parser_matrix = subparsers.add_parser(
        #     "matrix", help="cost v schedule chart. In development not working."
        # )
        # parser_data_query = subparsers.add_parser(
        #     "query", help="return data from core data"
        # )

        # parser_gmpp_ar = subparsers.add_parser(
        #     "gmpp_ar", help="compiled summaries for the IPA GMPP annual report"
        # )
        #
        # # Arguments
        # # stage
        # for sub in [
        #     parser_dca,
        #     parser_vfm,
        #     parser_risks,
        #     parser_port_risks,
        #     parser_speedial,
        #     # parser_dandelion,
        #     parser_costs,
        #     parser_costs_sp,
        #     # parser_data_query,
        #     parser_milestones,
        #     parser_data_query,
        # ]:
        #     sub.add_argument(
        #         "--stage",
        #         type=str,
        #         metavar="",
        #         action="store",
        #         nargs="*",
        #         choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
        #         help="Returns analysis for only those projects at the specified planning stage(s). By default "
        #         "the --stage argument will return the list of bc_stages specified in the config file."
        #         'Or user can enter one or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
        #     )
        # # stage dandelion only
        # parser_dandelion.add_argument(
        #     "--stage",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     nargs="*",
        #     choices=["FBC", "OBC", "SOBC", "pre-SOBC", "pipeline"],
        #     help="Returns analysis for only those projects at the specified planning stage(s). By default "
        #     "the --stage argument will return the list of bc_stages specified in the config file."
        #     'Or user can enter one or combination of "FBC", "OBC", "SOBC", "pre-SOBC".For dandelion "pipeline" is also available',
        # )
        #
        # # group
        # for sub in [
        #     parser_dca,
        #     parser_vfm,
        #     parser_risks,
        #     parser_port_risks,
        #     parser_speedial,
        #     parser_dandelion,
        #     parser_costs,
        #     parser_milestones,
        #     parser_summaries,
        #     parser_costs_sp,
        #     parser_data_query,
        # ]:
        #     sub.add_argument(
        #         "--group",
        #         type=str,
        #         metavar="",
        #         action="store",
        #         nargs="+",
        #         help="Returns analysis for specified project(s), only. User must enter one or a combination of "
        #         'DfT Group names; "HSRG", "RSS", "RIG", "AMIS","RPE", or the project(s) acronym or full name.',
        #     )

        # # quarter
        # for sub in [
        #     parser_dca,
        #     parser_vfm,
        #     parser_risks,
        #     parser_port_risks,
        #     parser_speedial,
        #     parser_dandelion,
        #     parser_costs,
        #     parser_costs_sp,
        #     parser_data_query,
        #     parser_milestones,
        # ]:
        #     sub.add_argument(
        #         "--quarter",
        #         type=str,
        #         metavar="",
        #         action="store",
        #         nargs="+",
        #         help="Returns analysis for one or combination of specified quarters. "
        #         'User must use correct format e.g "Q3 19/20"',
        #     )
        #
        # parser_costs.add_argument(
        #     "--baseline",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     nargs="+",
        #     choices=["current"],
        #     help="baseline option for costs refactored in Q1 21/22. Choose 'current' to return project "
        #     "reported bls as well as latest forecast profile",
        # )
        #
        # parser_milestones.add_argument(
        #     "--type",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     nargs="+",
        #     choices=["Approval", "Assurance", "Delivery"],
        #     help="Returns analysis for specified type of milestones.",
        # )
        #
        # parser_speedial.add_argument(
        #     "--conf_type",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=["sro", "finance", "benefits", "schedule", "resource"],
        #     help="Returns analysis for specified confidence types. options are"
        #     "'sro', 'finance', 'benefits', 'schedule', 'resource'."
        #     " As of Q2 20/21 it only provides a three rag dial.",
        # )
        #
        # for sub in [parser_milestones, parser_data_query]:
        #     sub.add_argument(
        #         "--koi",
        #         type=str,
        #         action="store",
        #         nargs="+",
        #         help="Returns the specified keys of interest (KOI).",
        #     )
        #
        # for sub in [parser_milestones, parser_data_query]:
        #     sub.add_argument(
        #         "--koi_fn",
        #         type=str,
        #         action="store",
        #         help="provide name of csv file contain key names",
        #     )
        #
        # parser_milestones.add_argument(
        #     "--dates",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     nargs=2,
        #     help="dates for analysis. Must provide start date and then end date in format e.g. '1/1/2021' '1/1/2022'.",
        # )
        #
        # parser_dandelion.add_argument(
        #     "--type",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=[
        #         "spent",
        #         "remaining",
        #         "benefits",
        #         "ps resource",
        #         "contract resource",
        #         "total resource",
        #         "funded resource",
        #     ],
        #     help="Provide the type of value to include in dandelion. Options are"
        #     ' "spent", "remaining", "benefits", "ps resource", "contract resource", "total resource", "funded resource".',
        # )
        #
        # parser_dandelion.add_argument(
        #     "--order_by",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=["schedule"],
        #     help="Specify how project circles should be ordered: 'schedule' only current"
        #     " option.",
        # )
        #
        # parser_summaries.add_argument(
        #     "--type",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=["long", "short"],
        #     help="Specify which form of report is required. Options are 'short' or 'long'",
        # )
        #
        # parser_costs_sp.add_argument(
        #     "--type",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=["cat"],
        #     help="Provide the type of value to include in dandelion. Options are"
        #     ' "cat".',
        # )
        #
        # parser_dandelion.add_argument(
        #     "--angles",
        #     type=int,
        #     metavar="",
        #     action="store",
        #     nargs="+",
        #     # choices=['sro', 'finance', 'benefits', 'schedule', 'resource'],
        #     help="Use can manually enter angles for group bubbles",
        # )
        #
        # parser_dandelion.add_argument(
        #     "--confidence",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=["sro", "finance", "benefits", "schedule", "resource"],
        #     help="specify the confidence type to displayed for each project.",
        # )
        #
        # parser_dandelion.add_argument(
        #     "--pc",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=["G", "A/G", "A", "A/R", "R"],
        #     help="specify the colour for the overall portfolio circle",
        # )
        #
        # parser_dandelion.add_argument(
        #     "--circle_colour",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=["No", "Yes"],
        #     help="specify whether to colour circles with DCA rating colours",
        # )
        #
        # parser_dandelion.add_argument(
        #     "--circle_edge",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     choices=["forward_look", "ipa"],
        #     help="specify whether to colour circle edge with SRO forward look rating. "
        #     "Options are 'forward_look' or 'ipa'.",
        # )
        #
        # # chart
        # for sub in [parser_dandelion, parser_costs, parser_costs_sp, parser_milestones]:
        #     sub.add_argument(
        #         "--chart",
        #         type=str,
        #         metavar="",
        #         action="store",
        #         choices=["show", "save"],
        #         help="options for building and saving graph output. Commands are 'show' or 'save' ",
        #     )
        #

        #
        # parser_milestones.add_argument(
        #     "--blue_line",
        #     type=str,
        #     metavar="",
        #     action="store",
        #     help="Insert blue line into chart to represent a date. "
        #     'Options are "today" "config_date" or a date in correct format e.g. "1/1/2021".',
        # )
        #
        # cli_args = parser.parse_args(sys.argv[2:])
        # settings_switch(cli_args, "ipdc")

    def cdg(self):
        run_parsers()


if __name__ == "__main__":
    main()
