import configparser
import csv
import json
from typing import List, Dict
from dateutil import parser

from analysis_engine.core_data import get_group_meta_data, get_dandelion_meta_data
from analysis_engine.error_msgs import config_issue


def set_default_args(op_args, **kwargs):
    if "group" not in op_args and "stage" not in op_args:
        op_args["group"] = kwargs["group"]
        if op_args['subparser_name'] == "dandelion":
            op_args["group"] = get_dandelion_meta_data(kwargs)['port_group']
            # thought a if statement might be required here, but working ok without.
            op_args["angles"] = get_dandelion_meta_data(kwargs)['angles']
    if "stage" in op_args:
        if op_args["stage"] == []:
            op_args["stage"] = kwargs["stage"]
    if "quarter" not in op_args:
        op_args["quarter"] = [kwargs["quarters"]]
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
        if type(op_args["koi"]) is list:  # test in cli
            op_args["key"] = op_args["koi"]
        else:
            op_args["key"] = [op_args["koi"]]
        return op_args
    else:
        return op_args


def convert_date(date_str: str):
    """
    When date converted into json file the dates take the standard python format
    year-month-day. This function converts format to year-day-month. This function is
    used when the MilestoneData class is created. Seems to be the best place to deploy.
    """
    try:
        return parser.parse(date_str)  # returns datetime
    except TypeError:  # for a different data value e.g integer.
        return date_str
    except ValueError:  # for string data that is not a date.
        return date_str
    # is a ParserError necessary here also?


def get_board_date(op_args):
    try:
        config_path = op_args["root_path"] + op_args["config"]
        config = configparser.ConfigParser()
        config.read(config_path)
        date_str = config["GLOBALS"]["date"]
        return parser.parse(date_str, dayfirst=True).date()
    except:
        config_issue()


def get_remove_income(op_args):
    if op_args["report"] == "ipdc":
        try:
            config_path = op_args["root_path"] + op_args["config"]
            config = configparser.ConfigParser()
            config.read(config_path)
            return config["COSTS"]["remove_income"]
        except:
            config_issue()  # this could be improved to be more specific
    else:
        return []


def get_integration_data(op_args):
    try:
        config_path = op_args["root_path"] + op_args["config"]
        config = configparser.ConfigParser()
        config.read(config_path)
        op_args["project_map_path"] = config["GMPP INTEGRATION"]["project_map"]
        op_args["gmpp_data_path"] = config["GMPP INTEGRATION"]["gmpp_data"]
        op_args["key_map_path"] = config["GMPP INTEGRATION"]["key_map"]
        return op_args
    except:
        config_issue()


def get_masters_to_merge(op_args):
    try:
        config_path = op_args["root_path"] + op_args["config"]
        config = configparser.ConfigParser()
        config.read(config_path)
        msts = json.loads(config.get("MERGE", "masters_list"))  # to return a list
        op_args["masters_list"] = msts
    except:
        config_issue()


# def get_remove_income_totals(
#     confi_path: Path,
# ) -> Dict:
#     # Returns a list of dft groups
#     try:
#         config = configparser.ConfigParser()
#         config.read(confi_path)
#         dict = {
#             "remove income from totals": config["COSTS"]["remove_income"],
#         }
#     except:
#         logger.critical(
#             "Configuration file issue. Please check remove_income list in the COST section"
#         )
#         sys.exit(1)
#
#     return dict

#
# def check_remove(op_args):  # subcommand arg
#     if "remove" in op_args:
#         from analysis_engine.data import CURRENT_LOG
#
#         for p in op_args["remove"]:
#             if p + " successfully removed from analysis." not in CURRENT_LOG:
#                 logger.warning(
#                     p + " not recognised and therefore not removed from analysis."
#                     ' Please make sure "remove" entry is correct.'
#                 )
