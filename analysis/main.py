"""
cli for analysis engine.
currently working on number of different subcommands
sub commands.. so far.
vfm: is the name of the command to runs analysis

Options... so far.
-group: an option for a particular dft group of projects. str. specific options.
-stage: an option for a group of projects at a particular business case stage. str. specific options.
-quarter: specifies the quarter(s) for analysis. at least one str.

-stage and -group cannot be entered at same time current. Can sort.

Next steps:
- explore possibility of there being a way to 'initiate' analysis engine so
master data is stored in memory and arguments run directly from it. rather
than having to convert excel ws into python dict each time. This would also be
a useful first step as lots of data checking is done as part of Master Class
creation.
- have cli so that it is analysis, rather than main.py
- packaged onto PyPI.

"""

import argparse


from data import (
    get_master_data,
    Master,
    get_project_information, VfMData, root_path, vfm_into_excel, MilestoneData, put_milestones_into_wb, Projects,
    run_p_reports
)


def vfm(args):
    print("compiling vfm analysis")
    m = Master(get_master_data(), get_project_information())
    vfm_m = VfMData(m)
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
    print("VfM analysis has been compiled. Enjoy!")


def milestones(args):
    print("compiling milestone analysis")
    m = Master(get_master_data(), get_project_information())
    projects = m.project_stage["Q2 20/21"]["FBC"] + m.project_stage["Q2 20/21"][
        "OBC"] + [Projects.hs2_2b]
    milestone_data = MilestoneData(m, projects)
    milestone_data.filter_chart_info(milestone_type=["Approval", "Delivery"])
    run = put_milestones_into_wb(milestone_data)
    run.save(root_path / "output/milestone_data_output_with_notes.xlsx")


def summaries(args):
    print("compiling summaries")
    m = Master(get_master_data(), get_project_information())
    run_p_reports(m, m.dft_groups["Q3 20/21"]["RDM"])


def main():
    parser = argparse.ArgumentParser(prog='engine',
        description='value for money analysis')
    subparsers = parser.add_subparsers()
    parser_vfm = subparsers.add_parser('vfm',
                                       help="vfm help")
    parser_milestones = subparsers.add_parser('milestones',
                                              help='milestone help')
    parser_summaries = subparsers.add_parser('summaries',
                                             help='summaries')
    parser_vfm.add_argument('-s',
                        '--stage',
                        type=str,
                        action='store',
                        nargs='+',
                        choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
                        help='Returns analysis for a group a projects at specified stage')
    parser_vfm.add_argument('-g',
                        '--group',
                        type=str,
                        action='store',
                        nargs='+',
                        choices=["HSMRPG", "AMIS", "Rail", "RDM"],
                        help='Returns analysis for specified DfT Group')
    parser_vfm.add_argument('-q',
                        '--quarters',
                        type=str,
                        action='store',
                        nargs='+',
                        help='Returns analysis for specified quarters. Must be in format e.g Q3 19/20')
    # parser_milestones.add_argument()
    parser_vfm.set_defaults(func=vfm)
    parser_milestones.set_defaults(func=milestones)
    parser_summaries.set_defaults(func=summaries)
    args = parser.parse_args()
    # print(vars(args))
    args.func(vars(args))


if __name__ == "__main__":
    main()
