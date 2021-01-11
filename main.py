"""
cml for analysis engine. structure of clm could be something like:

analysis_engine vfm [optional argument, quarters] [optional argument, groups] [optional argument, same]
analysis_engine initiate. **this would covert master data files into a Master object, for passing into arguments.
"""


import argparse
from vfm.vfm_analysis import compile_vfm_analysis
from project_analysis.p_reports import run_p_reports

from data_mgmt.data import (
    get_master_data,
    Master,
    get_project_information,
)


def main():
    parser = argparse.ArgumentParser(description='You are running analysis engine')
    parser.add_argument('-initiate',
                        metavar='run',
                        action='store',
                        help='stores all data on machine')
    parser.add_argument('-vfm',
                        metavar='run',
                        action='store',
                        help='runs vfm analysis')
    args = vars(parser.parse_args())
    print(args)
    if args["initiate"]:
        m = Master(get_master_data(), get_project_information())
    if args["vfm"]:
        print("compiling vfm analysis")
        compile_vfm_analysis(m)


if __name__ == "__main__":
    main()

# my_parser = argparse.ArgumentParser(description='Run analysis engine')
# my_parser.add_argument('vfm_analysis',
#                        action='store')
#                        # nargs='*',
#                        # default='vfm')
# my_parser.add_argument('project_sums',
#                        action='store')
#                        # default='run')
#
# args = my_parser.parse_args()
#
# if args.vfm_analysis == 'vfm':
#     print("compiling vfm analysis")
#     compile_vfm_analysis()
#
# if args.project_sums == 'run':
#     print("compiling project reports")
#     m = Master(get_master_data(), get_project_information())
#     projects = m.current_projects
#     run_p_reports(projects, m)