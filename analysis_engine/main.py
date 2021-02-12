"""
cli for analysis_engine engine.
currently working on number of different subcommands
sub commands.. so far.
vfm: is the name of the command to runs analysis_engine
summaries: in place but hard coded without options
milestones: in place but hard coded without options

Options... so far.
-group: an option for a particular dft group of projects. str. specific options.
-stage: an option for a group of projects at a particular business case stage. str. specific options.
-quarter: specifies the quarter(s) for analysis_engine. at least one str.

-stage and -group cannot be entered at same time current. Can sort.

Next steps:
- explore possibility of there being a way to 'initiate' analysis_engine engine so
master data is stored in memory and subcommands run directly from it. rather
than having to convert excel ws into python dict each time. This would also be
a useful first step as lots of data checking is done as part of Master Class
creation.
- have cli so that it is analysis_engine, rather than main.py
- packaged onto PyPI.

"""

import argparse
import itertools
import sys

from openpyxl import load_workbook

from analysis_engine.data import (
    get_master_data,
    Master,
    get_project_information,
    VfMData,
    root_path,
    vfm_into_excel,
    MilestoneData,
    put_milestones_into_wb,
    run_p_reports,
    RiskData,
    risks_into_excel,
    DcaData,
    dca_changes_into_excel,
    dca_changes_into_word,
    open_word_doc,
    Pickle,
    open_pickle_file,
    ipdc_dashboard,
    CostData,
    cost_v_schedule_chart_into_wb,
    make_file_friendly,
    DandelionData,
    dandelion_data_into_wb,
    run_dandelion_matplotlib_chart,
    put_matplotlib_fig_into_word,
    cost_profile_into_wb,
    cost_profile_graph,
    data_query_into_wb,
    get_data_query_key_names,
    ProjectNameError,
    milestone_chart, get_cost_stackplot_data, cost_stackplot_graph, cal_group,
)

import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s: %(levelname)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
)
logger = logging.getLogger(__name__)


def run_correct_args(
    m: Master,
    ae_class: MilestoneData or CostData or VfMData or DcaData or RiskData,
    args: argparse.ArgumentParser,
) -> MilestoneData or CostData or VfMData or DcaData:
    if args["quarters"] and args["stage"]:
        data = ae_class(m, quarter=args["quarters"], stage=args["stage"])
    elif args["quarters"] and args["group"]:
        data = ae_class(m, quarter=args["quarters"], group=args["group"])
    elif args["baselines"] and args["stage"]:
        data = ae_class(m, baseline=args["baselines"], stage=args["stage"])
    elif args["baselines"] and args["group"]:
        data = ae_class(m, baseline=args["baselines"], group=args["group"])
    elif args["baselines"]:
        data = ae_class(m, baseline=args["baselines"])
    elif args["quarters"]:
        data = ae_class(m, quarter=args["quarters"])
    elif args["stage"]:
        data = ae_class(m, quarter=["standard"], group=args["stage"])
    elif args["group"]:
        data = ae_class(m, quarter=["standard"], group=args["group"])
    else:
        data = ae_class(m, quarter=["standard"])

    return data


def get_args_for_file(args: argparse) -> list:
    l = []  # l is list
    for x in args.values():
        if x is not None:
            ffx = make_file_friendly(x)  # ffx
            l.append(ffx)
    l = l[1:-1]  # get rid of builtin_function_or_method
    unpack = itertools.chain.from_iterable(l)
    return list(unpack)


def initiate(args):
    print("creating a master data file for analysis_engine")
    try:
        master = Master(get_master_data(), get_project_information())
    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)

    path_str = str("{0}/core_data/pickle/master".format(root_path))
    Pickle(master, path_str)


def run_general(args):
    programme = args["subparser_name"]
    print("compiling " + programme + " analysis")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    try:
        if programme == "vfm":
            c = run_correct_args(m, VfMData, args)  # c is class
            wb = vfm_into_excel(c)
        if programme == "risks":
            c = run_correct_args(m, RiskData, args)
            wb = risks_into_excel(c)
        if programme == "dcas":
            c = run_correct_args(m, DcaData, args)
            wb = dca_changes_into_excel(c)
        if programme == "speedial":
            doc = open_word_doc(root_path / "input/summary_temp.docx")
            c = run_correct_args(m, DcaData, args)
            c.get_changes()
            doc = dca_changes_into_word(c, doc)
            doc.save(root_path / "output/{}.docx".format(programme))
            print(programme + " analysis has been compiled. Enjoy!")
        if programme == "dandelion":
            doc = open_word_doc(root_path / "input/summary_temp.docx")
            c = run_correct_args(m, DandelionData, args)
            wb = dandelion_data_into_wb(c)
            if args['chart']:
                for i in c.iter_list:
                    graph = run_dandelion_matplotlib_chart(c.d_data[i], chart=True)
                    if args['chart'] == 'save':
                        put_matplotlib_fig_into_word(doc, graph, size=6, transparent=True)
                        doc.save(root_path / "output/dandelion_chart.docx")
        if programme == "costs":
            doc = open_word_doc(root_path / "input/summary_temp.docx")
            c = run_correct_args(m, CostData, args)
            wb = cost_profile_into_wb(c)
            if args['chart']:
                if args["title"]:
                    graph = cost_profile_graph(c, title=args["title"], chart=True)
                else:
                    graph = cost_profile_graph(c, chart=True)
                if args['chart'] == 'save':
                    put_matplotlib_fig_into_word(doc, graph, size=6, transparent=False)
                    doc.save(root_path / "output/costs_chart.docx")

        if programme != "speedial":  # only excel outputs
            wb.save(root_path / "output/{}.xlsx".format(programme))
            print(programme + " analysis has been compiled. Enjoy!")
    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)

    # TODO optional_args produces a list of strings, each of which are to be in the output file name path.
    # optional_args = get_args_for_file(args)
    # wb.save(root_path / "output/{}_{}.xlsx".format(programme, optional_args))
    # print(programme + " analysis has been compiled. Enjoy!")


def milestones(args):
    print("compiling milestone analysis_engine")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    # print(args)
    try:
        # options "baselines" "quarters" "group" "dates" "stage"
        # bls
        if args["baselines"] and args["stage"] and args["dates"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["baselines"] and args["group"] and args["dates"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["baselines"] and args["dates"]:
            ms = MilestoneData(m, baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["baselines"] and args["group"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baselines"])
        elif args["baselines"]:
            ms = MilestoneData(m, baseline=args["baselines"])
        # quarters
        elif args["quarters"] and args["stage"] and args["dates"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["quarters"] and args["group"] and args["dates"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["quarters"] and args["group"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"])
        elif args["quarters"] and args["dates"]:
            ms = MilestoneData(m, quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["quarters"]:
            ms = MilestoneData(m, quarter=args["quarters"])
        # dates
        elif args["dates"] and args["group"]:
            ms = MilestoneData(m, quarter=["standard"], group=args["group"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["dates"]:
            ms = MilestoneData(m, quarter=["standard"])
            ms.filter_chart_info(dates=args["dates"])
        else:
            ms = MilestoneData(m, quarter=["standard"])

        wb = put_milestones_into_wb(ms)
        wb.save(root_path / "output/milestone_data_output.xlsx")

        if args['chart']:
            milestone_chart(ms, title="test", chart=True)
    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)
    except Warning:
        logger.critical("To many milestone for chart. Stopping")
        sys.exit(1)


def summaries(args):
    print("compiling summaries")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    if args["group"]:
        run_p_reports(m, group=args["group"], baseline=["standard"])
    else:
        run_p_reports(m, baseline=["standard"])


def dashboard(args):
    print("compiling ipdc dashboards")
    dashboard_master = load_workbook(root_path / "input/dashboards_master.xlsx")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    wb = ipdc_dashboard(m, dashboard_master)
    wb.save(root_path / "output/completed_ipdc_dashboard.xlsx")
    print("dashboard compiled. enjoy!")


def matrix(args):
    print("compiling cost and schedule matrix analysis")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    costs = CostData(m, m.current_projects)
    miles = MilestoneData(m, m.current_projects)
    miles.calculate_schedule_changes()
    wb = cost_v_schedule_chart_into_wb(miles, costs)
    wb.save(root_path / "output/costs_schedule_matrix.xlsx")
    print("Cost and schedule matrix compiled. Enjoy!")


def costs_sp(args):
    try:
        print("compiling cost stackplot")
        m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
        if args["group"]:
            g = cal_group(args["group"], m, 0)
            sp = get_cost_stackplot_data(m, g, "Q3 20/21", type="comp")
        else:
            sp = get_cost_stackplot_data(m, ['HSMRPG', 'Rail', 'RPE', 'AMIS'], "Q3 20/21", type="comp")
        cost_stackplot_graph(sp)
    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)

def query(args):
    print("Getting data")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    if args["keys"]:
        wb = data_query_into_wb(m, keys=args["keys"])
        wb.save(root_path / "output/data_query.xlsx")
        print("Data compiled. Enjoy!")
    elif args["file_name"] and args["quarters"]:
        l = get_data_query_key_names(
            root_path / "input/{}.csv".format(args["file_name"])
        )
        wb = data_query_into_wb(m, keys=l, quarters=args["quarters"])
        wb.save(root_path / "output/data_query.xlsx")
        print("Data compiled using " + args["file_name"] + ".cvs file. Enjoy!")
    else:
        l = get_data_query_key_names(root_path / "input/key_names.csv")
        wb = data_query_into_wb(m, keys=l)
        wb.save(root_path / "output/data_query.xlsx")
        print("Data compiled using key_names cvs file. Enjoy!")


def main():
    parser = argparse.ArgumentParser(
        prog="engine", description="DfT Major Projects Portfolio Office analysis engine"
    )
    subparsers = parser.add_subparsers(dest="subparser_name")
    parser_initiate = subparsers.add_parser(
        "initiate", help="creates a master data file"
    )
    parser_dashboard = subparsers.add_parser("dashboards", help="ipdc dashboard")
    parser_dandelion = subparsers.add_parser(
        "dandelion",
        help="dandelion graph and data (early version of graph output).",
    )
    parser_costs = subparsers.add_parser(
        "costs",
        help="cost trend profile graph and data (early version needs more testing).",
    )
    parser_costs_sp = subparsers.add_parser(
        "costs_sp",
        help="cost stackplot graph and data (early version needs more testing).",
    )
    parser_milestones = subparsers.add_parser(
        "milestones",
        help="milestone schedule graphs and data (early version needs more testing)",
    )
    parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
    parser_summaries = subparsers.add_parser("summaries", help="summary reports")
    parser_risks = subparsers.add_parser("risks", help="risk analysis")
    parser_dca = subparsers.add_parser("dcas", help="dca analysis")
    parser_speedial = subparsers.add_parser("speedial", help="speed dial analysis")
    parser_matrix = subparsers.add_parser("matrix", help="cost v schedule chart")
    parser_data_query = subparsers.add_parser(
        "query", help="return data from core data"
    )

    parser_summaries.add_argument(
        "--group",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        help="Returns summaries for specified project(s). User can either input DfT Group name; "
        '"HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym',
    )

    parser_costs_sp.add_argument(
        "--group",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        help="Returns summaries for specified project(s). User can either input DfT Group name; "
             '"HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym',
    )

    parser_data_query.add_argument(
        "--keys",
        type=str,
        metavar="Key Name",
        action="store",
        nargs="+",
        help="Returns the specified data keys.",
    )

    parser_data_query.add_argument(
        "--file_name",
        type=str,
        action="store",
        help="provide name of csc file contain key names",
    )

    parser_costs.add_argument(
        "--title", type=str, action="store", help="provide a title for chart. Optional"
    )

    parser_milestones.add_argument(
        "--baselines",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["current", "last", "bl_one", "bl_two", "bl_three", "standard", "all"],
        help="Returns analysis for specified baselines. Must be in correct format",
    )

    parser_speedial.add_argument(
        "--baselines",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["current", "last", "bl_one", "bl_two", "bl_three", "all"],
        help="Returns analysis for specified baselines. Must be in correct format",
    )

    parser_dca.add_argument(
        "--baselines",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["current", "last", "bl_one", "bl_two", "bl_three", "all"],
        help="Returns analysis for specified baselines. Must be in correct format",
    )

    parser_risks.add_argument(
        "--baselines",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["current", "last", "bl_one", "bl_two", "bl_three", "all"],
        help="Returns analysis for specified baselines. Must be in correct format",
    )

    parser_costs.add_argument(
        "--baselines",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["current", "last", "bl_one", "bl_two", "bl_three", "standard", "all"],
        help="Returns analysis for specified baselines. Must be in correct format",
    )

    parser_dandelion.add_argument(
        "--baselines",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["current", "last", "bl_one", "bl_two", "bl_three", "all"],
        help="Returns analysis for specified baselines. Must be in correct format",
    )

    parser_vfm.add_argument(
        "--baselines",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["current", "last", "bl_one", "bl_two", "bl_three", "all"],
        help="Returns analysis for specified baselines. Must be in correct format",
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
        "--chart",
        type=str,
        action="store",
        choices=['show', 'save'],
        help="options for building and saving graph output. Commands are 'show' or 'save' "
    )

    parser_costs.add_argument(
        "--chart",
        type=str,
        action="store",
        choices=['show', 'save'],
        help="options for building and saving graph output. Commands are 'show' or 'save' "
    )

    parser_milestones.add_argument(
        "--chart",
        type=str,
        action="store",
        choices=['show', 'save'],
        help="options for building and saving graph output. Commands are 'show' or 'save' "
    )

    for sub in [
        parser_dca,
        parser_vfm,
        parser_risks,
        parser_speedial,
        parser_dandelion,
        parser_costs,
        parser_data_query,
        parser_milestones,
    ]:
        # all sub-commands have the same optional args. This is working
        # but prob could be refactored.
        sub.add_argument(
            "--stage",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
            help="Returns analysis for those projects at the specified planning stage(s). Must be one "
            'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
        )
        sub.add_argument(
            "--group",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            # choices=["HSMRPG", "AMIS", "Rail", "RPE"],
            help="Returns summaries for specified project(s). User can either input DfT Group name; "
            '"HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym',
        )
        # no quarters in dandelion yet
        sub.add_argument(
            "--quarters",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            help="Returns analysis for specified quarters. Must be in format e.g Q3 19/20",
        )

    parser_initiate.set_defaults(func=initiate)
    parser_dashboard.set_defaults(func=dashboard)
    parser_dandelion.set_defaults(func=run_general)
    parser_costs.set_defaults(func=run_general)
    parser_vfm.set_defaults(func=run_general)
    parser_milestones.set_defaults(func=milestones)
    parser_summaries.set_defaults(func=summaries)
    parser_risks.set_defaults(func=run_general)
    parser_dca.set_defaults(func=run_general)
    parser_speedial.set_defaults(func=run_general)
    parser_matrix.set_defaults(func=matrix)
    parser_data_query.set_defaults(func=query)
    parser_costs_sp.set_defaults(func=costs_sp)
    args = parser.parse_args()
    # print(vars(args))
    args.func(vars(args))


if __name__ == "__main__":
    main()
