import argparse
from vfm.vfm_analysis import compile_vfm_analysis
from project_analysis.p_reports import run_p_reports

from data_mgmt.data import (
    get_master_data,
    Master,
    get_project_information,
)


my_parser = argparse.ArgumentParser(description='Run vfm analysis')
my_parser.add_argument('-a',
                       action='store',
                       # nargs='*',
                       default='run')
my_parser.add_argument('-b',
                       action='store',
                       default='run')

args = my_parser.parse_args()

if args.a == 'run':
    print("compiling vfm analysis")
    compile_vfm_analysis()

if args.b == 'run':
    print("compiling project reports")
    m = Master(get_master_data(), get_project_information())
    projects = m.current_projects
    run_p_reports(projects, m)