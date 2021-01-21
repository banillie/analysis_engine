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


from analysis_engine.data import (
    get_master_data,
    Master,
    get_project_information,
    VfMData,
    root_path,
    vfm_into_excel,
    MilestoneData,
    put_milestones_into_wb,
    Projects,
    run_p_reports,
    RiskData,
    risks_into_excel, DcaData, dca_changes_into_excel,
)


def vfm(args):
    print("compiling vfm analysis_engine")
    m = Master(get_master_data(), get_project_information())
    vfm_m = VfMData(
        m
    )  # why does this need to come first and not as else statement below?
    if args["quarters"]:
        vfm_m = VfMData(m, quarters=args["quarters"])
    if args["stage"]:
        vfm_m = VfMData(m, stage=args["stage"])
    if args["group"]:
        vfm_m = VfMData(m, group=args["group"])
    if args["quarters"] and args["stage"]:  # to test
        vfm_m = VfMData(m, quarters=args["quarters"], stage=args["stage"])
    if args["quarters"] and args["group"]:  # to test
        vfm_m = VfMData(m, quarters=args["quarters"], group=args["group"])

    wb = vfm_into_excel(vfm_m)
    wb.save(root_path / "output/vfm.xlsx")
    print("VfM analysis_engine has been compiled. Enjoy!")


def risks(args):
    print("compiling risk analysis_engine")
    m = Master(get_master_data(), get_project_information())
    risk_m = RiskData(
        m
    )  # why does this need to come first and not as else statement below?
    if args["quarters"]:
        risk_m = RiskData(m, quarters=args["quarters"])
    if args["stage"]:
        risk_m = RiskData(m, stage=args["stage"])
    if args["group"]:
        risk_m = RiskData(m, group=args["group"])
    if args["quarters"] and args["stage"]:  # to test
        risk_m = RiskData(m, quarters=args["quarters"], stage=args["stage"])
    if args["quarters"] and args["group"]:  # to test
        risk_m = RiskData(m, quarters=args["quarters"], group=args["group"])

    wb = risks_into_excel(risk_m)
    wb.save(root_path / "output/risks.xlsx")
    print("Risk analysis_engine has been compiled. Enjoy!")


def milestones(args):
    print("compiling milestone analysis_engine")
    m = Master(get_master_data(), get_project_information())
    projects = (
        m.project_stage["Q2 20/21"]["FBC"]
        + m.project_stage["Q2 20/21"]["OBC"]
        + [Projects.hs2_2b]
    )
    milestone_data = MilestoneData(m, projects)
    milestone_data.filter_chart_info(milestone_type=["Approval", "Delivery"])
    run = put_milestones_into_wb(milestone_data)
    run.save(root_path / "output/milestone_data_output_with_notes.xlsx")


def summaries(args):
    print("compiling summaries")
    proj_info = get_project_information()
    m = Master(get_master_data(), get_project_information())
    if args["group"]:
        run_p_reports(m, proj_info, group=args["group"])
    else:
        run_p_reports(m, proj_info)


def dca(args):
    m = Master(get_master_data(), get_project_information())
    dca = DcaData(m)
    # assert dca.dca_count == {}
    quarter_list = ["Q4 19/20", "Q4 18/19"]
    wb = dca_changes_into_excel(dca, quarter_list)
    wb.save(root_path / "output/dcas.xlsx")


def main():
    parser = argparse.ArgumentParser(
        prog="engine", description="DfT Major Projects Portfolio Office analysis engine"
    )
    subparsers = parser.add_subparsers()
    # subparsers.metavar = '                '
    parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
    parser_milestones = subparsers.add_parser("milestones", help="milestone analysis")
    parser_summaries = subparsers.add_parser("summaries", help="summary reports")
    parser_risks = subparsers.add_parser("risks", help="risk analysis")
    parser_dca = subparsers.add_parser("dcas", help="dca analysis")
    parser_vfm.add_argument(
        "--stage",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
        help="Returns analysis for those projects at the specified planning stage(s). Must be one "
        'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
    )
    parser_vfm.add_argument(
        "--group",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["HSMRPG", "AMIS", "Rail", "RDM"],
        help="Returns analysis for those projects in the specified DfT Group. Must be one or "
        'combination of "HSMRPG", "AMIS", "Rail", "RDM"',
    )
    parser_vfm.add_argument(
        "--quarters",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        help="Returns analysis for specified quarters. Must be in format e.g Q3 19/20",
    )
    parser_summaries.add_argument(
        "--group",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        # choices=["HSMRPG", "AMIS", "Rail", "RDM"],
        help="Returns summaries for specified projects. User can either input DfT Group name; "
        '"HSMRPG", "AMIS", "Rail", "RDM", or the project(s) acronym',
    )
    parser_risks.add_argument(
        "--stage",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
        help="Returns analysis for those projects at the specified planning stage(s). Must be one "
        'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
    )
    parser_risks.add_argument(
        "--group",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["HSMRPG", "AMIS", "Rail", "RDM"],
        help="Returns analysis for those projects in the specified DfT Group. Must be one or "
        'combination of "HSMRPG", "AMIS", "Rail", "RDM"',
    )
    parser_risks.add_argument(
        "--quarters",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        help="Returns analysis for specified quarters. Must be in format e.g Q3 19/20",
    )

    parser_vfm.set_defaults(func=vfm)
    parser_milestones.set_defaults(func=milestones)
    parser_summaries.set_defaults(func=summaries)
    parser_risks.set_defaults(func=risks)
    parser_dca.set_defaults(func=dca)
    args = parser.parse_args()
    # print(vars(args))
    args.func(vars(args))


if __name__ == "__main__":
    main()
