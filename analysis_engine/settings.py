import configparser
import platform
import csv
from pathlib import Path
from typing import List, Dict
from dateutil import parser

from datamaps.api import project_data_from_master, project_data_from_master_month

from analysis_engine.error_msgs import config_issue


def _platform_docs_dir(dir: str) -> Path:
    #  Cross plaform file path handling. The dir (directorary) controls the report type.
    if platform.system() == "Linux":
        return Path.home() / "Documents" / dir
    if platform.system() == "Darwin":
        return Path.home() / "Documents" / dir
    else:
        return Path.home() / "Documents" / dir


#
# INITIATE_DICT = {
#     'cdg': {
#         'report': 'cdg',
#         'root_path': str(_platform_docs_dir('cdg')),
#         'config': '/core_data/cdg_config.ini',
#         'callable': project_data_from_master
#     },
#     'ipdc': {
#         'config': '/core_data/ipdc_config.ini',
#         'callable': project_data_from_master,
#     },
#     'top_250': {
#         'config': '/core_data/top_250_config.ini',
#         'callable': project_data_from_master_month,
#     }
# }  # controls the documents pointed to for reporting process via cli positional arguments.
#


def report_config(report_type: str):
    if report_type == "cdg" or report_type == "ipdc":
        func = project_data_from_master
    if report_type == "top_250":
        func = project_data_from_master_month
    return {
        "report": report_type,
        "root_path": str(_platform_docs_dir(report_type)),
        "config": f"/core_data/{report_type}_config.ini",
        "callable": func,
        "master_path": "/core_data/json/master.json",
        "dashboard": "/input/dashboard_master.xlsx",
        "narrative_dashboard": "/input/narrative_dashboard_master.xlsx",
        "excel_save_path": "/output/{}.xlsx",
        "word_save_path": "/output/{}.docx",
        "word_landscape": "/input/summary_temp_landscape.docx",
        "word_portrait": "/input/summary_temp.docx",
    }
    # return INITIATE_DICT[report_type]


def set_default_args(op_args, port_group, default_quarter):
    if "group" not in op_args and "stage" not in op_args:
        op_args["group"] = port_group
    if "quarter" not in op_args:
        op_args["quarter"] = [default_quarter]
    if "chart" not in op_args:
        op_args["chart"] = False

    return op_args


def get_data_query_key_names(key_file: csv) -> List[str]:
    key_list = []
    with open(key_file) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=",")
        for row in csv_reader:
            key_list.append(row[0])
    return key_list[1:]


def return_koi_fn_keys(op_args: Dict):  # op_args
    """small helper function to convert key names in file into list of strings
    and place in op_args dictionary"""
    if "koi_fn" in op_args:
        keys = get_data_query_key_names(
            op_args["root_path"] + "/input/{}.csv".format(op_args["koi_fn"])
        )
        op_args["key"] = keys
        return op_args
    if "koi" in op_args:
        op_args["key"] = op_args["koi"]
        return op_args
    else:
        return op_args


def get_board_date(op_args):
    try:
        config_path = op_args["root_path"] + op_args["config"]
        config = configparser.ConfigParser()
        config.read(config_path)
        date_str = config["GLOBALS"]["milestones_blue_line_date"]
        return parser.parse(date_str).date()
    except:
        config_issue()
