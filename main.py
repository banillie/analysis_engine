"""
cli for analysis engine.
currently working on one example - vfm analysis

Arguments
vfm: is the name of the command to runs analysis
-group: an option for a particular dft group of projects. str. specific options.
-stage: an option for a group of projects at a particular business case stage. str. specific options.
-quarter: specifies the quarter(s) for analysis. at least one str.

-stage and -group cannot be entered at same time.

Once above established:
- build other cli arguments for other analysis engine outputs e.g. milestones.
- explore possibility of there being a way to 'initiate' analysis engine so
master data is stored in memory and arguments run directly from it. rather
than having to convert excel ws into python dict each time. This would also be
a useful first step as lots of data checking is done as part of Master Class
creation.
- have cli so that it is analysis_engine, rather than main.py

"""

import argparse


from data_mgmt.data import (
    get_master_data,
    Master,
    get_project_information, VfMData, root_path, vfm_into_excel
)


def main():
    parser = argparse.ArgumentParser(prog='vfm',
        description='value for money analysis')
    parser.add_argument('-vfm',
                        action='store_true',
                        help='runs vfm analysis')
    parser.add_argument('-stage',
                        type=str,
                        action='store',
                        choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
                        help='returns analysis for a group a projects at specified stage')
    parser.add_argument('-group',
                        type=str,
                        action='store',
                        choices=["HSMRPG", "AMIS", "Rail", "RDM"],
                        help='returns analysis for specified DfT Group')
    parser.add_argument('-quarters',
                        type=str,
                        action='store',
                        nargs='+',
                        help='returns analysis for specified quarters')
    args = vars(parser.parse_args())
    if args["vfm"]:
        print("compiling vfm analysis")
        m = Master(get_master_data(), get_project_information())
        current_quarter = str(m.master_data[0].quarter)
        last_quarter = str(m.master_data[1].quarter)
        quarter_list = [current_quarter, last_quarter]
        vfm = VfMData(m)
        vfm.get_dictionary()
        vfm.get_count()
        if args["stage"] in ["FBC", "OBC", "SOBC", "pre-SOBC"]:
            vfm.get_dictionary(stage=args["stage"])
            vfm.get_count()
        if args["group"] in ["HSMRPG", "AMIS", "Rail", "RDM"]:
            vfm.get_dictionary(group=args["group"])
            vfm.get_count()
        if args["quarters"]:
            quarter_list = args["quarters"]
        wb = vfm_into_excel(vfm, quarter_list)
        wb.save(root_path / "output/vfm.xlsx")
        print("vfm analysis compiled")


if __name__ == "__main__":
    main()

