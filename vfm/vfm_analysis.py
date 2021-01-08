"""
Compiles VfM analysis. Places into output file an excel file with data.

This programme takes the project data from the following keys:
- state when code finalised.
In the output workbook the initial tabs contain a raw print out of this data for each
project. The final tab contains a 'count' as required for analysis.
What is the analysis trying to achieve... need from user(s).

Command line options
analysis_engine run vfm analysis. Default is current and latest quarter, all projects.
flags.
- quarters. specify the quarters to be analysed
- groups. specify particular groups e.g. business case stage, dft group
- project. specify a particular project or group of projects.

"""
import argparse
import sys

from data_mgmt.data import (
    Master,
    root_path,
    get_project_information,
    get_master_data,
    VfMData,
    vfm_into_excel
)


def compile_vfm_analysis():
    m = Master(get_master_data(), get_project_information())
    vfm = VfMData(m)
    latest_quarter = str(m.master_data[0].quarter)
    last_quarter = str(m.master_data[1].quarter)
    default_quarter_list = [latest_quarter, last_quarter]
    wb = vfm_into_excel(vfm, default_quarter_list)
    wb.save(root_path / "output/vfm.xlsx")


my_parser = argparse.ArgumentParser(description='Run vfm analysis')
my_parser.add_argument('input',
                       action='store',
                       nargs='*',
                       default='run')

args = my_parser.parse_args()

if args == "run":
    compile_vfm_analysis()
