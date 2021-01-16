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
- have cli so that it is analysis_engine, rather than main.py
- packaged onto PyPI.

"""

import argparse


from analysis_engine.data import (
    get_master_data,
    Master,
    get_project_information, VfMData, root_path, vfm_into_excel
)


def vfm(args):
    print("compiling vfm analysis")
    m = Master(get_master_data(), get_project_information())
    vfm_m = VfMData(m)
    if args["quarters"]:
        vfm_m = VfMData(m, quarters=args["quarters"])
    if args["stage"][0] in ["FBC", "OBC", "SOBC", "pre-SOBC"]:
        vfm_m = VfMData(m, stage=args["stage"])
    if args["group"] in ["HSMRPG", "AMIS", "Rail", "RDM"]:
        vfm_m = VfMData(m, group=args["group"])
    if args["quarters"] and args["stage"][0] in ["FBC", "OBC", "SOBC", "pre-SOBC"]:
        vfm_m = VfMData(m, quarters=args["quarters"], stage=args["stage"])
    if args["quarters"] and args["group"] in ["HSMRPG", "AMIS", "Rail", "RDM"]:
        vfm_m = VfMData(m, quarters=args["quarters"], group=args["group"])
    wb = vfm_into_excel(vfm_m)
    wb.save(root_path / "output/vfm.xlsx")
    print("VfM analysis has been compiled. Enjoy!")


def main():
    parser = argparse.ArgumentParser(prog='engine',
        description='value for money analysis')
    subparsers = parser.add_subparsers()
    parser_vfm = subparsers.add_parser('vfm',
                                       help="vfm help")
    parser_vfm.add_argument('-s',
                        '--stage',
                        type=str,
                        action='store',
                        nargs='+',
                        choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
                        help='Returns analysis for a group a projects at specified stage')
    parser_vfm.add_argument('-group',
                        type=str,
                        action='store',
                        nargs='+',
                        choices=["HSMRPG", "AMIS", "Rail", "RDM"],
                        help='Returns analysis for specified DfT Group')
    parser_vfm.add_argument('-quarters',
                        type=str,
                        action='store',
                        nargs='+',
                        help='Returns analysis for specified quarters. Must be in format e.g Q3 19/20')
    parser_vfm.set_defaults(func=vfm)
    args = parser.parse_args()
    args.func(vars(args))


if __name__ == "__main__":
    main()

