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
    put_matplotlib_fig_into_word, cost_profile_into_wb, cost_profile_graph,
)


def run_correct_args(
    m: Master,
    ae_class: MilestoneData or CostData or VfMData or DcaData or RiskData,
    args: argparse.ArgumentParser,
) -> MilestoneData or CostData or VfMData or DcaData:
    # try:  # acts as partition for subcommand options
    if args["quarters"] and args["stage"]:  # to test
        data = ae_class(m, quarters=args["quarters"], stage=args["stage"])
    elif args["quarters"] and args["group"]:  # to test
        data = ae_class(m, quarters=args["quarters"], group=args["group"])
    elif args["quarters"]:
        data = ae_class(m, quarters=args["quarters"])
    elif args["stage"]:
        data = ae_class(m, stage=args["stage"])
    elif args["group"]:
        data = ae_class(m, group=args["group"])
    else:
        data = ae_class(m)
    # except KeyError:
    #     data = ae_class(m)

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
    master = Master(get_master_data(), get_project_information())
    path_str = str("{0}/core_data/pickle/master".format(root_path))
    Pickle(master, path_str)


def run_general(args):
    programme = args["subparser_name"]
    print("compiling " + programme + " analysis")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
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
        report_doc = open_word_doc(root_path / "input/summary_temp.docx")
        c = run_correct_args(m, DcaData, args)
        c.get_changes()
        doc = dca_changes_into_word(c, report_doc)
        doc.save(root_path / "output/{}.docx".format(programme))
        print(programme + " analysis has been compiled. Enjoy!")
    if programme == "dandelion":
        report_doc = open_word_doc(root_path / "input/summary_temp.docx")
        c = run_correct_args(m, DandelionData, args)
        wb = dandelion_data_into_wb(c)
        graph = run_dandelion_matplotlib_chart(c)
        put_matplotlib_fig_into_word(report_doc, graph, size=4, transparent=True)
        report_doc.save(root_path / "output/dandelion_output.docx")
    if programme == "costs":
        c = run_correct_args(m, CostData, args)
        wb = cost_profile_into_wb(c)
        cost_profile_graph(c)

    if programme != "speedial":  # only excel outputs
        wb.save(root_path / "output/{}.xlsx".format(programme))
        print(programme + " analysis has been compiled. Enjoy!")

    # TODO optional_args produces a list of strings, each of which are to be in the output file name path.
    # optional_args = get_args_for_file(args)
    # wb.save(root_path / "output/{}_{}.xlsx".format(programme, optional_args))
    # print(programme + " analysis has been compiled. Enjoy!")


def milestones(args):
    print("compiling milestone analysis_engine")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    # projects = (
    #     m.project_stage["Q2 20/21"]["FBC"]
    #     + m.project_stage["Q2 20/21"]["OBC"]
    #     + [Projects.hs2_2b]
    # )
    milestone_data = MilestoneData(m, m.current_projects)
    milestone_data.filter_chart_info(milestone_type=["Approval", "Delivery"])
    run = put_milestones_into_wb(milestone_data)
    run.save(root_path / "output/milestone_data_output_with_notes_q3.xlsx")


def summaries(args):
    print("compiling summaries")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    if args["group"]:
        run_p_reports(m, m.project_information, group=args["group"])
    else:
        run_p_reports(m, m.project_information)


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
        "dandelion", help="dandelion graph (early version) and data"
    )
    parser_costs = subparsers.add_parser("costs", help="cost analysis")
    parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
    parser_milestones = subparsers.add_parser("milestones", help="milestone analysis")
    parser_summaries = subparsers.add_parser("summaries", help="summary reports")
    parser_risks = subparsers.add_parser("risks", help="risk analysis")
    parser_dca = subparsers.add_parser("dcas", help="dca analysis")
    parser_speedial = subparsers.add_parser("speedial", help="speed dial analysis")
    parser_matrix = subparsers.add_parser("matrix", help="cost v schedule chart")

    # parser_vfm.add_argument(
    #     "--stage",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
    #     help="Returns analysis for those projects at the specified planning stage(s). Must be one "
    #     'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
    # )
    # parser_vfm.add_argument(
    #     "--group",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     choices=["HSMRPG", "AMIS", "Rail", "RDM"],
    #     help="Returns analysis for those projects in the specified DfT Group. Must be one or "
    #     'combination of "HSMRPG", "AMIS", "Rail", "RDM"',
    # )
    # parser_vfm.add_argument(
    #     "--quarters",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     help="Returns analysis for specified quarters. Must be in format e.g Q3 19/20",
    # )

    # parser_risks.add_argument(
    #     "--stage",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
    #     help="Returns analysis for those projects at the specified planning stage(s). Must be one "
    #     'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
    # )
    # parser_risks.add_argument(
    #     "--group",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     choices=["HSMRPG", "AMIS", "Rail", "RDM"],
    #     help="Returns analysis for those projects in the specified DfT Group. Must be one or "
    #     'combination of "HSMRPG", "AMIS", "Rail", "RDM"',
    # )
    # parser_risks.add_argument(
    #     "--quarters",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     help="Returns analysis for specified quarters. Must be in format e.g Q3 19/20",
    # )
    # parser_dca.add_argument(
    #     "--stage",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
    #     help="Returns analysis for those projects at the specified planning stage(s). Must be one "
    #          'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
    # )
    # parser_dca.add_argument(
    #     "--group",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     choices=["HSMRPG", "AMIS", "Rail", "RDM"],
    #     help="Returns analysis for those projects in the specified DfT Group. Must be one or "
    #          'combination of "HSMRPG", "AMIS", "Rail", "RDM"',
    # )
    # parser_dca.add_argument(
    #     "--quarters",
    #     type=str,
    #     metavar="",
    #     action="store",
    #     nargs="+",
    #     help="Returns analysis for specified quarters. Must be in format e.g Q3 19/20",
    # )

    parser_summaries.add_argument(
        "--group",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        help="Returns summaries for specified projects. User can either input DfT Group name; "
        '"HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym',
    )

    for sub in [
        parser_dca,
        parser_vfm,
        parser_risks,
        parser_speedial,
        parser_dandelion,
        parser_costs
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
            choices=["HSMRPG", "AMIS", "Rail", "RPE"],
            help="Returns analysis for those projects in the specified DfT Group. Must be one or "
            'combination of "HSMRPG", "AMIS", "Rail", "RPE"',
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
    args = parser.parse_args()
    # print(vars(args))
    args.func(vars(args))


if __name__ == "__main__":
    main()
