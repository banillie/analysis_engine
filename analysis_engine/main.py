import argparse
from argparse import RawTextHelpFormatter
import sys
from typing import Dict, List
from pathlib import Path
import configparser
import json

from analysis_engine import __version__

from datamaps.api import project_data_from_master, project_data_from_master_month

from analysis_engine.data import (
    get_master_data,
    Master,
    get_project_information,
    VfMData,
    root_path,
    cdg_root_path,
    vfm_into_excel,
    MilestoneData,
    put_milestones_into_wb,
    run_p_reports,
    RiskData,
    risks_into_excel,
    DcaData,
    dca_changes_into_excel,
    dca_changes_into_word,
    Pickle,
    open_pickle_file,
    ipdc_dashboard,
    CostData,
    cost_v_schedule_chart_into_wb,
    DandelionData,
    put_matplotlib_fig_into_word,
    cost_profile_into_wb,
    cost_profile_graph,
    data_query_into_wb,
    get_data_query_key_names,
    ProjectNameError,
    ProjectGroupError,
    ProjectStageError,
    milestone_chart,
    cost_stackplot_graph,
    make_a_dandelion_auto,
    build_speedials,
    get_sp_data,
    # DFT_GROUP,
    get_input_doc,
    InputError,
    JsonMaster,
    JsonData,
    open_json_file,
    cost_profile_into_wb_new,
    cost_profile_graph_new,
    get_gmpp_data,
    place_gmpp_online_keys_into_dft_master_format,
    data_check_print_out,
)

import logging

from analysis_engine.top35_data import (
    top35_run_p_reports,
    top35_root_path,
    CentralSupportData,
)

from analysis_engine.cdg_data import (
    cdg_dashboard,
    cdg_narrative_dashboard,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s: %(levelname)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
)
logger = logging.getLogger(__name__)

class ConfigurationError(Exception):
    pass

def get_group_stage_data(
    confi_path: Path,
) -> List[str]:
    # Returns a list of dft groups
    try:
        config = configparser.ConfigParser()
        config.read(confi_path)
        master_data_list = []
        portfolio_group = json.loads(
            config.get("GROUPS", "portfolio_groups")
        )  # to return a list
        group_all = json.loads(config.get("GROUPS", "all_groups"))
        try:
            bc_stages = json.loads(config.get("GROUPS", "bc_stages"))
        except configparser.NoOptionError:
            bc_stages = []
    except:
        logger.critical("Configuration file issue. Please check and make sure it's correct.")
        sys.exit(1)

    return portfolio_group, group_all, bc_stages


def check_remove(op_args):  # subcommand arg
    if "remove" in op_args:
        from analysis_engine.data import CURRENT_LOG

        for p in op_args["remove"]:
            if p + " successfully removed from analysis." not in CURRENT_LOG:
                logger.warning(
                    p + " not recognised and therefore not removed from analysis."
                    ' Please make sure "remove" entry is correct.'
                )


def ipdc_initiate(args):
    print("creating a master data file for ipdc and gmpp portfolio reporting.")

    # get group information from config. only all_groups used at initiate
    META = get_group_stage_data(
        str(root_path) + "/core_data/ipdc_config.ini",
    )
    all_groups = META[1]

    try:
        master = JsonMaster(
            get_master_data(
                str(root_path) + "/core_data/ipdc_config.ini",
                str(root_path) + "/core_data/",
                project_data_from_master,
            ),
            get_project_information(
                str(root_path) + "/core_data/ipdc_config.ini",
                str(root_path) + "/core_data/",
            ),
            all_groups,
        )
        master.get_baseline_data()
        master.check_baselines()
    except (ProjectNameError, ProjectGroupError, ProjectStageError) as e:
        logger.critical(e)
        sys.exit(1)

    master_json_path = str("{0}/core_data/json/master".format(root_path))
    JsonData(master, master_json_path)


def top250_initiate(args):
    print("creating a master data file for top 250 reporting.")

    # get group information from config. only all_groups used at initiate
    META = get_group_stage_data(
        str(top35_root_path) + "/core_data/top250_config.ini",
    )
    all_groups = META[1]

    try:
        master = JsonMaster(
            get_master_data(
                str(top35_root_path) + "/core_data/top250_config.ini",
                str(top35_root_path) + "/core_data/",
                project_data_from_master_month,
            ),
            get_project_information(
                str(top35_root_path) + "/core_data/top250_config.ini",
                str(top35_root_path) + "/core_data/",
            ),
            all_groups,
            data_type="top35",
        )
    except (ProjectNameError, ProjectGroupError, ProjectStageError) as e:
        logger.critical(e)
        sys.exit(1)

    master_json_path = str("{0}/core_data/json/master".format(top35_root_path))
    JsonData(master, master_json_path)


def cdg_initiate(args):
    print("creating a master data file for CDG reporting.")

    # get group information from config. only all_groups used at initiate
    META = get_group_stage_data(
        str(cdg_root_path) + "/core_data/cdg_config.ini",
    )
    all_groups = META[1]

    try:
        master = JsonMaster(
            get_master_data(
                str(cdg_root_path) + "/core_data/cdg_config.ini",
                str(cdg_root_path) + "/core_data/",
                project_data_from_master,
            ),
            get_project_information(
                str(cdg_root_path) + "/core_data/cdg_config.ini",
                str(cdg_root_path) + "/core_data/",
            ),
            all_groups,
            data_type="cdg",
        )
    except (ProjectNameError, ProjectGroupError, ProjectStageError) as e:
        logger.critical(e)
        sys.exit(1)

    master_json_path = str("{0}/core_data/json/master".format(cdg_root_path))
    JsonData(master, master_json_path)


def ipdc_run_general(args):
    # get portfolio reporting group information.
    META = get_group_stage_data(
        str(root_path) + "/core_data/ipdc_config.ini",
    )
    dft_group = META[0]
    dft_stage = META[2]

    programme = args["subparser_name"]
    # wrap this into logging
    try:
        print("compiling ipdc " + programme + " analysis")
    except TypeError:  # NoneType as no programme entered
        print("Further command required. Use --help flag for guidance")
        sys.exit(1)

    m = Master(open_json_file(str(root_path / "core_data/json/master.json")))

    try:
        # print(args.items())
        op_args = {
            k: v for k, v in args.items() if v is not None
        }  # removes None values
        # print(op_args)
        if "group" not in op_args:
            if "stage" not in op_args:
                op_args["group"] = dft_group
            if "stage" in op_args:
                if op_args["stage"] == []:
                    op_args["stage"] = dft_stage
        if "quarter" not in op_args:
            if "baseline" not in op_args:
                op_args["quarter"] = ["standard"]

        # print(op_args)

        if programme == "vfm":
            c = VfMData(m, **op_args)  # c is class
            wb = vfm_into_excel(c)

        if programme == "risks":
            c = RiskData(m, **op_args)
            wb = risks_into_excel(c)

        if programme == "dcas":
            c = DcaData(m, **op_args)
            wb = dca_changes_into_excel(c)

        if programme == "costs":
            if "baseline" in op_args:
                if op_args["baseline"] == ["current"]:
                    op_args["quarter"] = [str(m.current_quarter)]
                    c = CostData(m, **op_args)
                    c.get_forecast_cost_profile()
                    c.get_baseline_cost_profile()
            else:
                c = CostData(m, **op_args)
                c.get_forecast_cost_profile()
            wb = cost_profile_into_wb_new(c)
            if "chart" not in op_args:
                op_args["chart"] = True
                cost_profile_graph_new(c, m, **op_args)
            else:
                if op_args["chart"] == "save":
                    op_args["chart"] = False
                    cost_graph = cost_profile_graph_new(c, m, **op_args)
                    doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
                    put_matplotlib_fig_into_word(doc, cost_graph, size=7.5)
                    doc.save(root_path / "output/costs_graph.docx")
                if op_args["chart"] == "show":
                    op_args["chart"] = True
                    cost_profile_graph_new(c, m, **op_args)

        if programme == "costs_sp":
            sp_data = get_sp_data(m, **op_args)

            if "chart" not in op_args:
                op_args["chart"] = True
                cost_stackplot_graph(sp_data, m, **op_args)
            else:
                if op_args["chart"] == "save":
                    op_args["chart"] = False
                    sp_graph = cost_stackplot_graph(sp_data, m, **op_args)
                    doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
                    put_matplotlib_fig_into_word(doc, sp_graph, size=7.5)
                    doc.save(root_path / "output/stack_plot_graph.docx")
                if op_args["chart"] == "show":
                    op_args["chart"] = True
                    cost_stackplot_graph(sp_data, m, **op_args)

        if programme == "speedial":
            doc = get_input_doc(root_path / "input/summary_temp.docx")
            land_doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
            # if "conf_type" in op_args:
            #     if op_args["conf_type"] == "sro_three":
            #         op_args["rag_number"] = "3"
            #         op_args["quarter"] = [str(m.current_quarter)]
            #         data = DcaData(m, **op_args)
            #         build_speedials(data, land_doc)
            #         land_doc.save(root_path / "output/speed_dial_graph.docx")
            #     else:  # refactor!!
            op_args["rag_number"] = "3"
            data = DcaData(m, **op_args)
            data.get_changes()
            doc = dca_changes_into_word(data, doc)
            doc.save(root_path / "output/speed_dials_text.docx")
            build_speedials(data, land_doc)
            land_doc.save(root_path / "output/speed_dial_graph.docx")
            # else:
            #     op_args["rag_number"] = "5"
            #     data = DcaData(m, **op_args)
            #     data.get_changes()
            #     doc = dca_changes_into_word(data, doc)
            #     doc.save(root_path / "output/speed_dials_text.docx")
            #     build_speedials(data, land_doc)
            #     land_doc.save(root_path / "output/speed_dial_graph.docx")

        if programme == "milestones":
            ms = MilestoneData(m, **op_args)

            if (
                "type" in op_args
                or "dates" in op_args
                or "koi" in op_args
                or "koi_fn" in op_args
            ):
                op_args = return_koi_fn_keys(op_args)
                ms.filter_chart_info(**op_args)

            if "chart" not in op_args:
                pass
            else:
                if op_args["chart"] == "save":
                    op_args["chart"] = False
                    ms_graph = milestone_chart(ms, m, **op_args)
                    doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
                    put_matplotlib_fig_into_word(
                        doc, ms_graph, size=8, transparent=False
                    )
                    doc.save(root_path / "output/milestones_chart.docx")
                if op_args["chart"] == "show":
                    milestone_chart(ms, m, **op_args)

            wb = put_milestones_into_wb(ms)

        if programme == "dandelion":
            if op_args["quarter"] == [
                "standard"
            ]:  # converts "standard" default to current quarter
                op_args["quarter"] = [str(m.current_quarter)]
            # op_args["order_by"] = "schedule"
            d_data = DandelionData(m, **op_args)
            if "chart" not in op_args:
                op_args["chart"] = True
                make_a_dandelion_auto(d_data, **op_args)
            else:
                if op_args["chart"] == "save":
                    op_args["chart"] = False
                    d_graph = make_a_dandelion_auto(d_data, **op_args)
                    doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
                    put_matplotlib_fig_into_word(doc, d_graph, size=7)
                    doc.save(root_path / "output/dandelion_graph.docx")
                if op_args["chart"] == "show":
                    make_a_dandelion_auto(d_data, **op_args)

        if programme == "dashboards":
            # op_args["baseline"] = ["standard"]
            dashboard_master = get_input_doc(root_path / "input/dashboards_master.xlsx")
            wb = ipdc_dashboard(m, dashboard_master, op_args)
            wb.save(root_path / "output/completed_ipdc_dashboard.xlsx")

        if programme == "summaries":
            op_args["quarter"] = [str(m.current_quarter)]
            if "type" not in op_args:
                op_args["type"] = "short"
            run_p_reports(m, **op_args)

        # if programme == "top_250_summaries":
        #     if op_args["group"] == dft_group:
        #         op_args["group"] = ["HSRG", "RSS", "RIG", "RPE"]
        #     top35_run_p_reports(m, **op_args)

        if programme == "matrix":
            costs = CostData(m, **op_args)
            miles = MilestoneData(m, *op_args)
            miles.calculate_schedule_changes()
            wb = cost_v_schedule_chart_into_wb(miles, costs)
            wb.save(root_path / "output/costs_schedule_matrix.xlsx")

        if programme == "query":
            if "koi" not in op_args and "koi_fn" not in op_args:
                logger.critical(
                    "Please enter a key name(s) using either --koi or --koi_fn"
                )
                sys.exit(1)
            op_args = return_koi_fn_keys(op_args)
            wb = data_query_into_wb(m, **op_args)

        if programme == "gmpp_integration":
            print("Coming soon!")
            # d = get_gmpp_data("GMPP_DATA_DFT_Q2_2021")  # gmpp_data
            # # place_gmpp_online_keys_into_dft_master_format(d, "KEY_MAP_v12", "master_2_2021")  # key_map, ipdc_master
            # p_list = ['East Coast Mainline Programme', 'Rapid Charging Fund', 'NO2 Reduction']  # list
            # place_gmpp_online_keys_into_dft_master_format(d, "KEY_MAP_v12", "master_2_2021", project_list=p_list)
            # # data_check_print_out("master_2_2021", "KEY_MAP_v12", "PROJECT_MAP")  # project map

        check_remove(op_args)

        try:
            if programme != "dashboards":
                wb.save(root_path / "output/{}.xlsx".format(programme))
        except UnboundLocalError:
            pass

        print(programme + " analysis has been compiled. Enjoy!")

    except (ProjectNameError, FileNotFoundError, InputError) as e:
        logger.critical(e)
        sys.exit(1)

    # TODO optional_args produces a list of strings, each of which are to be in the output file name path.
    # optional_args = get_args_for_file(args)
    # wb.save(root_path / "output/{}_{}.xlsx".format(programme, optional_args))
    # print(programme + " analysis has been compiled. Enjoy!")


def top250_run_general(args):
    programme = args["subparser_name"]
    print("compiling top250 " + programme + " analysis")
    m = Master(open_json_file(str(top35_root_path / "core_data/json/master.json")))

    # get portfolio reporting group information.
    META = get_group_stage_data(
        str(top35_root_path) + "/core_data/top250_config.ini",
    )
    dft_group = META[0]
    dft_stage = META[2]

    try:
        op_args = {k: v for k, v in args.items() if v}  # removes None values
        if "group" not in op_args:
            if "stage" not in op_args:
                op_args["group"] = dft_group
        if "quarter" not in op_args:
            if "baseline" not in op_args:
                op_args["quarter"] = ["standard"]

        op_args["data_type"] = "top35"

        if programme == "summaries":
            top35_run_p_reports(m, **op_args)

        if programme == "milestones" or programme == "central_support":
            if programme == "milestones":
                d = MilestoneData(m, **op_args)  # data
            if programme == "central_support":
                d = CentralSupportData(m, **op_args)

            if (
                "type" in op_args
                or "dates" in op_args
                or "koi" in op_args
                or "koi_fn" in op_args
            ):
                op_args = return_koi_fn_keys(op_args)
                d.filter_chart_info(**op_args)

            if "chart" not in op_args:
                pass
            else:
                if op_args["chart"] == "save":
                    op_args["chart"] = False
                    ms_graph = milestone_chart(d, m, **op_args)
                    doc = get_input_doc(
                        top35_root_path / "input/summary_temp_landscape.docx"
                    )
                    put_matplotlib_fig_into_word(
                        doc, ms_graph, size=8, transparent=False
                    )
                    doc.save(top35_root_path / "output/milestones_chart.docx")
                if op_args["chart"] == "show":
                    milestone_chart(d, m, **op_args)

            wb = put_milestones_into_wb(d)

        if programme == "dandelion":
            if op_args["quarter"] == ["standard"]:
                # converts "standard" default to current quarter
                op_args["quarter"] = [str(m.current_quarter)]
            d_data = DandelionData(m, **op_args)
            if "chart" not in op_args:
                op_args["chart"] = True
                make_a_dandelion_auto(d_data, **op_args)
            else:
                if op_args["chart"] == "save":
                    op_args["chart"] = False
                    d_graph = make_a_dandelion_auto(d_data, **op_args)
                    doc = get_input_doc(
                        top35_root_path / "input/summary_temp_landscape.docx"
                    )
                    put_matplotlib_fig_into_word(doc, d_graph, size=7)
                    doc.save(top35_root_path / "output/dandelion_graph.docx")
                if op_args["chart"] == "show":
                    make_a_dandelion_auto(d_data, **op_args)

        if programme == "query":
            if "koi" not in op_args and "koi_fn" not in op_args:
                logger.critical(
                    "Please enter a key name(s) using either --keys or --file_name"
                )
                sys.exit(1)
            op_args = return_koi_fn_keys(op_args)
            wb = data_query_into_wb(m, **op_args)

        check_remove(op_args)

        try:
            if programme != "dashboards":
                wb.save(top35_root_path / "output/{}.xlsx".format(programme))
        except UnboundLocalError:
            pass

        print(programme + " analysis has been compiled. Enjoy!")

    except (ProjectNameError, FileNotFoundError, InputError) as e:
        logger.critical(e)
        sys.exit(1)


def cdg_run_general(args):
    # get portfolio reporting group information.
    META = get_group_stage_data(
        str(cdg_root_path) + "/core_data/cdg_config.ini",
    )
    dft_group = META[0]
    dft_stage = META[2]

    programme = args["subparser_name"]
    # wrap this into logging
    try:
        print("compiling cdg " + programme + " analysis")
    except TypeError:  # NoneType as no programme entered
        print("Further command required. Use --help flag for guidance")
        sys.exit(1)

    m = Master(open_json_file(str(cdg_root_path / "core_data/json/master.json")))

    try:
        op_args = {
            k: v for k, v in args.items() if v is not None
        }  # removes None values

        op_args["data_type"] = "cdg"

        if "group" not in op_args:
            if "stage" not in op_args:
                op_args["group"] = dft_group
            if "stage" in op_args:
                if op_args["stage"] == []:
                    op_args["stage"] = dft_stage
        if "quarter" not in op_args:
            if "baseline" not in op_args:
                op_args["quarter"] = ["standard"]

        if programme == "speedial":
            doc = get_input_doc(cdg_root_path / "input/summary_temp.docx")
            land_doc = get_input_doc(cdg_root_path / "input/summary_temp_landscape.docx")
            op_args["rag_number"] = "5"
            data = DcaData(m, **op_args)
            data.get_changes()
            doc = dca_changes_into_word(data, doc)
            doc.save(cdg_root_path / "output/speed_dials_text.docx")
            build_speedials(data, land_doc)
            land_doc.save(cdg_root_path / "output/speed_dial_graph.docx")

        if programme == "milestones":
            print(op_args)
            ms = MilestoneData(m, **op_args)

            if (
                "type" in op_args
                or "dates" in op_args
                or "koi" in op_args
                or "koi_fn" in op_args
            ):
                op_args = return_koi_fn_keys(op_args)
                ms.filter_chart_info(**op_args)

            if "chart" not in op_args:
                pass
            else:
                if op_args["chart"] == "save":
                    op_args["chart"] = False
                    ms_graph = milestone_chart(ms, m, **op_args)
                    doc = get_input_doc(root_path / "input/summary_temp_landscape.docx")
                    put_matplotlib_fig_into_word(
                        doc, ms_graph, size=8, transparent=False
                    )
                    doc.save(root_path / "output/milestones_chart.docx")
                if op_args["chart"] == "show":
                    milestone_chart(ms, m, **op_args)

            wb = put_milestones_into_wb(ms)

        if programme == "dandelion":
            if op_args["quarter"] == [
                "standard"
            ]:  # converts "standard" default to current quarter
                op_args["quarter"] = [str(m.current_quarter)]
            # op_args["order_by"] = "schedule"
            d_data = DandelionData(m, **op_args)
            if "chart" not in op_args:
                op_args["chart"] = True
                make_a_dandelion_auto(d_data, **op_args)
            else:
                if op_args["chart"] == "save":
                    op_args["chart"] = False
                    d_graph = make_a_dandelion_auto(d_data, **op_args)
                    doc = get_input_doc(cdg_root_path / "input/summary_temp_landscape.docx")
                    put_matplotlib_fig_into_word(doc, d_graph, size=7)
                    doc.save(cdg_root_path / "output/dandelion_graph.docx")
                if op_args["chart"] == "show":
                    make_a_dandelion_auto(d_data, **op_args)

        if programme == "dashboards":
            dashboard_master = get_input_doc(cdg_root_path / "input/dashboard_master.xlsx")
            narrative_d_master = get_input_doc(cdg_root_path / "input/narrative_dashboard_master.xlsx")
            wb = cdg_narrative_dashboard(m, narrative_d_master)
            wb.save(cdg_root_path / "output/{}.xlsx".format("cdg_narrative_dashboard_completed"))
            wb = cdg_dashboard(m, dashboard_master)
            wb.save(cdg_root_path / "output/{}.xlsx".format("cdg_dashboard_completed"))

        check_remove(op_args)

        try:
            if programme != "dashboards":
                wb.save(cdg_root_path / "output/{}.xlsx".format(programme))
        except UnboundLocalError:
            pass

        print(programme + " analysis has been compiled. Enjoy!")

    except (ProjectNameError, FileNotFoundError, InputError) as e:
        logger.critical(e)
        sys.exit(1)


def return_koi_fn_keys(oa: Dict):  # op_args
    """small helper function to convert key names in file into list of strings
    and place in op_args dictionary"""
    if "koi_fn" in oa:
        if "data_type" in oa:
            if oa["data_type"] == "top35":
                keys = get_data_query_key_names(
                    top35_root_path / "input/{}.csv".format(oa["koi_fn"])
                )
        else:
            keys = get_data_query_key_names(
                root_path / "input/{}.csv".format(oa["koi_fn"])
            )
        oa["key"] = keys
        return oa
    if "koi" in oa:
        oa["key"] = oa["koi"]
        return oa
    else:
        return oa


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
            help="Initial command to specify whether ipdc, top250, cdg analysis if required",
        )
        # parse_args defaults to [1:] for args, but you need to
        # exclude the rest of the args too, or validation will fail
        args = parser.parse_args(sys.argv[1:2])
        if vars(args)["command"] not in ["ipdc", "top250", "cdg"]:
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
        parser_risks = subparsers.add_parser("risks", help="risk analysis")
        parser_dca = subparsers.add_parser("dcas", help="dca analysis")
        parser_speedial = subparsers.add_parser("speedial", help="speed dial analysis")
        parser_matrix = subparsers.add_parser(
            "matrix", help="cost v schedule chart. In development not working."
        )
        parser_data_query = subparsers.add_parser(
            "query", help="return data from core data"
        )
        parser_gmpp_integration = subparsers.add_parser(
            "gmpp_integration", help="integrates gmpp data"
        )

        # Arguments
        # stage
        for sub in [
            parser_dca,
            parser_vfm,
            parser_risks,
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
            parser_speedial,
            parser_dandelion,
            parser_costs,
            # parser_data_query,
            parser_milestones,
            parser_summaries,
            parser_costs_sp,
            parser_data_query,
            # parser_top_250_summaries,
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
        # baseline
        for sub in [
            parser_dca,
            parser_vfm,
            parser_risks,
            parser_speedial,
            # parser_dandelion,
            # parser_costs,
            parser_data_query,
            parser_milestones,
        ]:
            sub.add_argument(
                "--baseline",
                type=str,
                metavar="",
                action="store",
                nargs="+",
                choices=[
                    "current",
                    "last",
                    "bl_one",
                    "bl_two",
                    "bl_three",
                    "standard",
                    "all",
                ],
                help="Returns analysis for specified baselines. User must use correct format"
                ' which are "current", "last", "bl_one", "bl_two", "bl_three", "standard", "all".'
                ' The "all" option returns all, "standard" returns first three',
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
            choices=["spent", "remaining", "benefits"],
            help="Provide the type of value to include in dandelion. Options are"
            ' "spent", "remaining", "benefits".',
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
            'Options are "Today" "IPDC" or a date in correct format e.g. "1/1/2021".',
        )
        # now that we're inside a subcommand, ignore the first
        # TWO argvs, ie the command (git) and subcommand (commit)
        # print(sys.argv)
        args = parser.parse_args(sys.argv[2:])
        # print(args)
        # print(vars(args))
        # print(vars(args))
        if vars(args)["subparser_name"] == "initiate":
            ipdc_initiate(vars(args))
        else:
            ipdc_run_general(vars(args))

    def top250(self):
        parser = argparse.ArgumentParser(
            description="runs all analysis for top250 reporting"
        )
        subparsers = parser.add_subparsers(dest="subparser_name")
        subparsers.metavar = "                      "
        top250_parser_initiate = subparsers.add_parser(
            "initiate", help="creates a master data file"
        )
        top250_parser_summaries = subparsers.add_parser(
            "summaries", help="summary reports"
        )
        top250_parser_milestones = subparsers.add_parser(
            "milestones",
            help="milestone schedule graphs and data.",
        )
        top250_parser_cs = subparsers.add_parser(
            "central_support",
            help="central support schedule graphs and data.",
        )
        dandelion_description = (
            "Creates the 'dandelion' graph. See below optional arguments for changing the "
            "dandelion that is compiled. The command analysis dandelion returns the default "
            'dandelion graph. The user must specify --chart "save" to save the chart, otherwise '
            "only a temporary matplotlib chart will be generated."
        )
        top250_parser_dandelion = subparsers.add_parser(
            "dandelion",
            help="Dandelion graph.",
            description=dandelion_description,
            # formatter_class=RawTextHelpFormatter,  # can't use as effects how optional arguments are shown.
        )
        top250_parser_data_query = subparsers.add_parser(
            "query", help="return data from core data"
        )

        for sub in [
            top250_parser_milestones,
            top250_parser_summaries,
            top250_parser_cs,
            top250_parser_dandelion,
            top250_parser_data_query,
        ]:
            sub.add_argument(
                "--group",
                type=str,
                metavar="",
                action="store",
                nargs="+",
                help="Returns analysis for specified project(s), only. User must enter one or a combination of "
                'DfT Group names; "HSRG", "RSS", "RIG", "RPE", or the project(s) acronym or full name.',
            )

        for sub in [
            top250_parser_milestones,
            top250_parser_cs,
            top250_parser_dandelion,
            top250_parser_data_query,
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

        for sub in [
            top250_parser_milestones,
            top250_parser_cs,
        ]:
            sub.add_argument(
                "--dates",
                type=str,
                metavar="",
                action="store",
                nargs=2,
                help="dates for analysis. Must provide start date and then end date in format e.g. '1/1/2021' '1/1/2022'.",
            )

        for sub in [
            top250_parser_milestones,
            top250_parser_cs,
            top250_parser_dandelion,
        ]:
            sub.add_argument(
                "--chart",
                type=str,
                metavar="",
                action="store",
                choices=["show", "save"],
                help="options for building and saving graph output. Commands are 'show' or 'save' ",
            )

        for sub in [
            top250_parser_milestones,
            top250_parser_data_query,
        ]:
            sub.add_argument(
                "--koi",
                type=str,
                action="store",
                nargs="+",
                help="Returns the specified keys of interest (KOI).",
            )

        for sub in [
            top250_parser_milestones,
            top250_parser_cs,
        ]:
            sub.add_argument(
                "--blue_line",
                type=str,
                metavar="",
                action="store",
                help="Insert blue line into chart to represent a date. "
                'Options are "Today" "CDG" or a date in correct format e.g. "1/1/2021".',
            )

        for sub in [
            top250_parser_milestones,
            top250_parser_data_query,
        ]:
            sub.add_argument(
                "--koi_fn",
                type=str,
                action="store",
                help="provide name of csc file contain key names",
            )

        top250_parser_dandelion.add_argument(
            "--angles",
            type=int,
            metavar="",
            action="store",
            nargs="+",
            # choices=['sro', 'finance', 'benefits', 'schedule', 'resource'],
            help="Use can manually enter angles for group bubbles",
        )

        args = parser.parse_args(sys.argv[2:])
        # print(vars(args))
        if vars(args)["subparser_name"] == "initiate":
            top250_initiate(vars(args))
        else:
            top250_run_general(vars(args))

    def cdg(self):
        parser = argparse.ArgumentParser(
            description="runs all analysis for cdg reporting"
        )
        subparsers = parser.add_subparsers(dest="subparser_name")
        subparsers.metavar = "                      "
        cdg_parser_initiate = subparsers.add_parser(
            "initiate", help="creates a master data file"
        )
        cdg_parser_milestones = subparsers.add_parser(
            "milestones",
            help="milestone schedule graphs and data.",
        )
        dandelion_description = (
            "Creates the 'dandelion' graph. See below optional arguments for changing the "
            "dandelion that is compiled. The command analysis dandelion returns the default "
            'dandelion graph. The user must specify --chart "save" to save the chart, otherwise '
            "only a temporary matplotlib chart will be generated."
        )
        cdg_parser_dandelion = subparsers.add_parser(
            "dandelion",
            help="Dandelion graph.",
            description=dandelion_description,
        )
        dashboard_description = (
            "Creates CDG dashboards. There are no optional arguments for this command.\n\n"
            "A blank master dashboard titled dashboards_master.xlsx must be in input file.\n\n"
            "A completed dashboard title completed_cdg_dashboard.xlsx will be placed into\n"
            "the output file."
        )
        cdg_parser_dashboard = subparsers.add_parser(
            "dashboards",
            help="CDG dashboard",
            description=dashboard_description,
            formatter_class=RawTextHelpFormatter,
        )

        cdg_parser_speedial = subparsers.add_parser("speedial", help="speed dial analysis")

        cdg_parser_milestones.add_argument(
                "--dates",
                type=str,
                metavar="",
                action="store",
                nargs=2,
                help="dates for analysis. Must provide start date and then end date in format e.g. '1/1/2021' '1/1/2022'.",
            )

        cdg_parser_milestones.add_argument(
            "--blue_line",
            type=str,
            metavar="",
            action="store",
            help="Insert blue line into chart to represent a date. "
                 'Options are "Today" "CDG" or a date in correct format e.g. "1/1/2021".',
        )

        for sub in [
            cdg_parser_milestones,
            cdg_parser_dandelion,
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
            cdg_parser_speedial,
            cdg_parser_dandelion,
            cdg_parser_milestones,
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

        cdg_parser_dandelion.add_argument(
            "--angles",
            type=int,
            metavar="",
            action="store",
            nargs="+",
            # choices=['sro', 'finance', 'benefits', 'schedule', 'resource'],
            help="Use can manually enter angles for group bubbles",
        )

        cdg_parser_dandelion.add_argument(
            "--type",
            type=str,
            metavar="",
            action="store",
            choices=["benefits", 'income'],
            help="Provide the type of value to include in dandelion. Options are"
                 ' "benefits" or "income".',
        )

        args = parser.parse_args(sys.argv[2:])
        # print(vars(args))
        if vars(args)["subparser_name"] == "initiate":
            cdg_initiate(vars(args))
        else:
            cdg_run_general(vars(args))


## old method for handling argparse commands.
# def main():
#     ae_description = (
#         "Welcome to the DfT Major Projects Portfolio Office analysis engine.\n\n"
#         "To operate use subcommands outlined below. To navigate each subcommand\n"
#         "option use the --help flag which will provide instructions on which optional\n"
#         "arguments can be used with each subcommand. e.g. analysis dandelion --help."
#     )
#     parser = argparse.ArgumentParser(
#         # prog="engine",
#         description=ae_description,
#         formatter_class=RawTextHelpFormatter
#     )
#     parser.add_argument('--version', action='version', version="0.0.22")  # link to setup.py
#     subparsers = parser.add_subparsers(dest="subparser_name")
#     subparsers.metavar = "                      "
#     # parser.add_argument('initiate', metavar="initiate", type=str, nargs=1, help="creates a master data file")
#     parser_initiate = subparsers.add_parser(
#         "initiate", help="creates a master data file"
#     )
#     dashboard_description = (
#         "Creates IPDC dashboards. There are no optional arguments for this command.\n\n"
#         "A blank master dashboard titled dashboards_master.xlsx must be in input file.\n\n"
#         "A completed dashboard title completed_ipdc_dashboard.xlsx will be placed into\n"
#         "the output file."
#     )
#     parser_dashboard = subparsers.add_parser(
#         "dashboards",
#         help="IPDC dashboards",
#         description=dashboard_description,
#         formatter_class=RawTextHelpFormatter,
#     )
#     dandelion_description = (
#         "Creates the IPDC 'dandelion' graph. See below optional arguments for changing the "
#         "dandelion that is compiled. The command analysis dandelion returns the default "
#         'dandelion graph. The user must specify --chart "save" to save the chart, otherwise '
#         "only a temporary matplotlib chart will be generated."
#     )
#     parser_dandelion = subparsers.add_parser(
#         "dandelion",
#         help="Dandelion graph.",
#         description=dandelion_description,
#         # formatter_class=RawTextHelpFormatter,  # can't use as effects how optional arguments are shown.
#     )
#
#     costs_description = (
#         "Creates a cost profile graph. See below optional arguments. The user "
#         'must specify --chart "save" to save the chart, otherwise '
#         "only a temporary matplotlib chart will be generated."
#     )
#
#     parser_costs = subparsers.add_parser(
#         "costs",
#         help="cost trend profile graph and data.",
#         description=costs_description,
#     )
#
#     costs_sp_description = (
#         "Creates a cost stack plot profile graph. See below optional arguments. The user "
#         'must specify --chart "save" to save the chart, otherwise '
#         "only a temporary matplotlib chart will be generated."
#     )
#
#     parser_costs_sp = subparsers.add_parser(
#         "costs_sp",
#         help="cost stack plot graph and data.",
#         description=costs_sp_description,
#     )
#
#     parser_milestones = subparsers.add_parser(
#         "milestones",
#         help="milestone schedule graphs and data.",
#     )
#     parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
#     parser_summaries = subparsers.add_parser("summaries", help="summary reports")
#     parser_risks = subparsers.add_parser("risks", help="risk analysis")
#     parser_dca = subparsers.add_parser("dcas", help="dca analysis")
#     parser_speedial = subparsers.add_parser("speedial", help="speed dial analysis")
#     parser_matrix = subparsers.add_parser(
#         "matrix", help="cost v schedule chart. In development not working."
#     )
#     parser_data_query = subparsers.add_parser(
#         "query", help="return data from core data"
#     )
#     parser_top_250_summaries = subparsers.add_parser(
#         "top_250_summaries", help="top 250 summaries"
#     )
#
#     # Arguments
#     # stage
#     for sub in [
#         parser_dca,
#         parser_vfm,
#         parser_risks,
#         parser_speedial,
#         # parser_dandelion,
#         parser_costs,
#         parser_costs_sp,
#         # parser_data_query,
#         parser_milestones,
#         parser_data_query,
#     ]:
#         sub.add_argument(
#             "--stage",
#             type=str,
#             metavar="",
#             action="store",
#             nargs="+",
#             choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
#             help="Returns analysis for only those projects at the specified planning stage(s). User must enter one "
#             'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
#         )
#     # stage dandelion only
#     parser_dandelion.add_argument(
#         "--stage",
#         type=str,
#         metavar="",
#         action="store",
#         nargs="+",
#         choices=["FBC", "OBC", "SOBC", "pre-SOBC", "pipeline"],
#         help="Returns analysis for only those projects at the specified planning stage(s). User must enter one "
#         'or combination of "FBC", "OBC", "SOBC", "pre-SOBC". For dandelion "pipeline" is also available',
#     )
#
#     # group
#     for sub in [
#         parser_dca,
#         parser_vfm,
#         parser_risks,
#         parser_speedial,
#         parser_dandelion,
#         parser_costs,
#         # parser_data_query,
#         parser_milestones,
#         parser_summaries,
#         parser_costs_sp,
#         parser_data_query,
#         parser_top_250_summaries,
#     ]:
#         sub.add_argument(
#             "--group",
#             type=str,
#             metavar="",
#             action="store",
#             nargs="+",
#             help="Returns analysis for specified project(s), only. User must enter one or a combination of "
#             'DfT Group names; "HSRG", "RSS", "RIG", "AMIS","RPE", or the project(s) acronym or full name.',
#         )
#     # remove
#     for sub in [
#         parser_dca,
#         parser_vfm,
#         parser_risks,
#         parser_speedial,
#         parser_dandelion,
#         parser_costs,
#         parser_costs_sp,
#         parser_data_query,
#         parser_milestones,
#     ]:
#         sub.add_argument(
#             "--remove",
#             type=str,
#             metavar="",
#             action="store",
#             nargs="+",
#             help="Removes specified projects from analysis. User must enter one or a combination of either"
#             " a recognised DfT Group name, a recognised planning stage or the project(s) acronym or full"
#             " name.",
#         )
#     # quarter
#     for sub in [
#         parser_dca,
#         parser_vfm,
#         parser_risks,
#         parser_speedial,
#         parser_dandelion,
#         parser_costs,
#         parser_costs_sp,
#         parser_data_query,
#         parser_milestones,
#     ]:
#         sub.add_argument(
#             "--quarter",
#             type=str,
#             metavar="",
#             action="store",
#             nargs="+",
#             help="Returns analysis for one or combination of specified quarters. "
#             'User must use correct format e.g "Q3 19/20"',
#         )
#     # baseline
#     for sub in [
#         parser_dca,
#         parser_vfm,
#         parser_risks,
#         parser_speedial,
#         # parser_dandelion,
#         parser_costs,
#         parser_data_query,
#         parser_milestones,
#     ]:
#         sub.add_argument(
#             "--baseline",
#             type=str,
#             metavar="",
#             action="store",
#             nargs="+",
#             choices=[
#                 "current",
#                 "last",
#                 "bl_one",
#                 "bl_two",
#                 "bl_three",
#                 "standard",
#                 "all",
#             ],
#             help="Returns analysis for specified baselines. User must use correct format"
#             ' which are "current", "last", "bl_one", "bl_two", "bl_three", "standard", "all".'
#             ' The "all" option returns all, "standard" returns first three',
#         )
#
#     parser_milestones.add_argument(
#         "--type",
#         type=str,
#         metavar="",
#         action="store",
#         nargs="+",
#         choices=["Approval", "Assurance", "Delivery"],
#         help="Returns analysis for specified type of milestones.",
#     )
#
#     for sub in [parser_milestones, parser_data_query]:
#         sub.add_argument(
#             "--koi",
#             type=str,
#             action="store",
#             nargs="+",
#             help="Returns the specified keys of interest (KOI).",
#         )
#
#     for sub in [parser_milestones, parser_data_query]:
#         sub.add_argument(
#             "--koi_fn",
#             type=str,
#             action="store",
#             help="provide name of csc file contain key names",
#         )
#
#     parser_milestones.add_argument(
#         "--dates",
#         type=str,
#         metavar="",
#         action="store",
#         nargs=2,
#         help="dates for analysis. Must provide start date and then end date in format e.g. '1/1/2021' '1/1/2022'.",
#     )
#
#     parser_dandelion.add_argument(
#         "--type",
#         type=str,
#         metavar="",
#         action="store",
#         choices=["spent", "remaining", "benefits"],
#         help="Provide the type of value to include in dandelion. Options are"
#         ' "spent", "remaining", "benefits".',
#     )
#
#     parser_costs_sp.add_argument(
#         "--type",
#         type=str,
#         metavar="",
#         action="store",
#         choices=["cat"],
#         help="Provide the type of value to include in dandelion. Options are" ' "cat".',
#     )
#
#     parser_dandelion.add_argument(
#         "--angles",
#         type=int,
#         metavar="",
#         action="store",
#         nargs="+",
#         # choices=['sro', 'finance', 'benefits', 'schedule', 'resource'],
#         help="Use can manually enter angles for group bubbles",
#     )
#
#     parser_dandelion.add_argument(
#         "--confidence",
#         type=str,
#         metavar="",
#         action="store",
#         choices=["sro", "finance", "benefits", "schedule", "resource"],
#         help="specify the confidence type to displayed for each project.",
#     )
#
#     parser_dandelion.add_argument(
#         "--pc",
#         type=str,
#         metavar="",
#         action="store",
#         choices=["G", "A/G", "A", "A/R", "R"],
#         help="specify the colour for the overall portfolio circle",
#     )
#
#     parser_dandelion.add_argument(
#         "--circle_colour",
#         type=str,
#         metavar="",
#         action="store",
#         choices=["No", "Yes"],
#         help="specify whether to colour circles with DCA rating colours",
#     )
#
#     # chart
#     for sub in [parser_dandelion, parser_costs, parser_costs_sp, parser_milestones]:
#         sub.add_argument(
#             "--chart",
#             type=str,
#             metavar="",
#             action="store",
#             choices=["show", "save"],
#             help="options for building and saving graph output. Commands are 'show' or 'save' ",
#         )
#
#     # title
#     for sub in [parser_costs, parser_milestones, parser_costs_sp]:
#         sub.add_argument(
#             "--title",
#             type=str,
#             metavar="",
#             action="store",
#             help="provide a title for chart. Optional",
#         )
#
#     parser_milestones.add_argument(
#         "--blue_line",
#         type=str,
#         metavar="",
#         action="store",
#         help="Insert blue line into chart to represent a date. "
#         'Options are "Today" "IPDC" or a date in correct format e.g. "1/1/2021".',
#     )
#
#     # parser_data_query.add_argument(
#     #     "--file_name",
#     #     type=str,
#     #     action="store",
#     #     help="provide name of csc file contain key names",
#     # )
#
#     # for sub in [
#     #     parser_dca,
#     #     parser_vfm,
#     #     parser_risks,
#     #     parser_speedial,
#     #     parser_dandelion,
#     #     parser_costs,
#     #     parser_data_query,
#     #     parser_milestones,
#     # ]:
#     #     # all sub-commands have the same optional args. This is working
#     #     # but prob could be refactored.
#     #     sub.add_argument(
#     #         "--stage",
#     #         type=str,
#     #         metavar="",
#     #         action="store",
#     #         nargs="+",
#     #         choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
#     #         help="Returns analysis for those projects at the specified planning stage(s). Must be one "
#     #         'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
#     #     )
#     #     sub.add_argument(
#     #         "--group",
#     #         type=str,
#     #         metavar="",
#     #         action="store",
#     #         nargs="+",
#     #         # choices=DFT_GROUP
#     #         ,
#     #         help="Returns summaries for specified project(s). User can either input DfT Group name; "
#     #         '"HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym',
#     #     )
#     #     # no quarters in dandelion yet
#     #     sub.add_argument(
#     #         "--quarters",
#     #         type=str,
#     #         metavar="",
#     #         action="store",
#     #         nargs="+",
#     #         help="Returns analysis for specified quarters. Must be in format e.g Q3 19/20",
#     #     )
#     #     sub.add_argument(
#     #         "--remove",
#     #         type=str,
#     #         metavar="",
#     #         action="store",
#     #         nargs="+",
#     #         # choices=DFT_GROUP
#     #         ,
#     #         help="Removes specified projects from analysis. User can either input DfT Group name; "
#     #              '"HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym"',
#     #     )
#
#     parser_initiate.set_defaults(func=initiate)
#     parser_dashboard.set_defaults(func=run_general)
#     parser_dandelion.set_defaults(func=run_general)
#     parser_costs.set_defaults(func=run_general)
#     parser_vfm.set_defaults(func=run_general)
#     # parser_milestones.set_defaults(func=run_general)
#     parser_summaries.set_defaults(func=run_general)
#     parser_risks.set_defaults(func=run_general)
#     parser_dca.set_defaults(func=run_general)
#     parser_speedial.set_defaults(func=run_general)
#     parser_matrix.set_defaults(func=run_general)
#     parser_data_query.set_defaults(func=run_general)
#     parser_costs_sp.set_defaults(func=run_general)
#     parser_top_250_summaries.set_defaults(func=run_general)
#     args = parser.parse_args()
#     # print(vars(args))
#     args.func(vars(args))


if __name__ == "__main__":
    main()
