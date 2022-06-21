import configparser
import json
import platform
import sys
import datetime
from pathlib import Path
from typing import List, Dict, Union, Optional, Tuple, TextIO, Callable
# from datetime import timedelta, date
from collections import OrderedDict

from openpyxl import load_workbook

from datamaps.process import Cleanser

from analysis_engine.error_msgs import (
    ProjectNameError, ProjectGroupError, ProjectStageError, logger
)


def _platform_docs_dir(dir: str) -> Path:
    #  Cross plaform file path handling. The dir (directorary) controls the report type.
    if platform.system() == "Linux":
        return Path.home() / "Documents" / dir
    if platform.system() == "Darwin":
        return Path.home() / "Documents" / dir
    else:
        return Path.home() / "Documents" / dir


root_path = _platform_docs_dir('ipdc')
cdg_root_path = _platform_docs_dir('cdg')


def get_master_data(
        confi_path: Path,
        pi_path: Path,
        func: Callable,
) -> List[Dict[str, Union[str, int, datetime.date, float]]]:
    """Returns a list of dictionaries each containing quarter data"""
    config = configparser.ConfigParser()
    config.read(confi_path)
    master_data_list = []
    for key in config["MASTERS"]:
        text = config["MASTERS"][key].split(", ")
        year = text[2]
        quarter = text[1]
        m_path = str(pi_path) + text[0]
        m = func(m_path, int(quarter), int(year))
        master_data_list.append(m)

    return list(reversed(master_data_list))


def convert_none_types(x):
    if x is None:
        return 0
    else:
        return x

class JsonMaster:
    def __init__(
            self,
            master_data: List[Dict[str, Union[str, int, datetime.date, float]]],
            project_information: Dict[str, Union[str, int]],
            all_groups,
            **kwargs,
    ) -> None:
        self.master_data = master_data
        self.project_information = project_information
        self.all_groups = all_groups
        self.all_projects = list(project_information.keys())
        self.kwargs = kwargs
        self.current_quarter = str(master_data[0].quarter)
        self.current_projects = master_data[0].projects
        self.abbreviations = {}
        self.full_names = {}
        self.bl_info = {}
        self.bl_index = {}
        self.dft_groups = {}
        self.project_group = {}
        self.project_stage = {}
        self.pipeline_dict = {}
        self.pipeline_list = []
        self.quarter_list = []
        self.get_quarter_list()
        # self.get_baseline_data()
        self.check_project_information()
        self.get_project_abbreviations()
        # self.check_baselines()
        self.get_project_groups()
        self.pipeline_projects_information()
        self.get_current_tp()

    def get_project_abbreviations(self) -> None:
        """gets the abbreviations for all current projects.
        held in the project info document"""
        abb_dict = {}
        fn_dict = {}
        error_case = []
        for p in self.all_projects:
            abb = self.project_information[p]["Abbreviations"]
            abb_dict[p] = {"abb": abb, "full name": p}
            fn_dict[abb] = p
            if abb is None:
                error_case.append(p)

        if error_case:
            for p in error_case:
                logger.critical("No abbreviation provided for " + p + ".")
            raise ProjectNameError(
                "Abbreviations must be provided for all projects in project_info. Program stopping. Please amend"
            )

        self.abbreviations = abb_dict
        self.full_names = fn_dict

    def get_baseline_data(self) -> None:
        """
        Returns two dictionaries used to calculate baselines.
        The current method is that each projects baseline list
        is prefixed with current and last quarter index (including None
        if necessary), as these need to be present at later stage
        calculations. structure of output dict is
        {bl_index: {bl_type: {proj_name: [baseline index list]}}}
        """

        # handles baselines across different datasets.
        if "data_type" in self.kwargs:
            if self.kwargs["data_type"] == "cdg":
                baseline_dict = CDG_BASELINE_TYPES
        else:
            baseline_dict = BASELINE_TYPES

        baseline_info = {}
        baseline_index = {}
        for b_type in list(baseline_dict.keys()):
            project_baseline_info = {}
            project_baseline_index = {}
            for name in self.current_projects:
                lower_list = []
                for i, master in list(enumerate(self.master_data)):
                    quarter = str(master.quarter)
                    if name in master.projects:
                        approved_bc = master.data[name][b_type]
                        if approved_bc == "Yes":
                            lower_list.append((approved_bc, quarter, i))
                    else:
                        pass

                if name in self.master_data[1].projects:  # prefix for other bl data
                    index_list = [0, 1]
                else:  # project not present last quarter so none
                    index_list = [0, None]
                for x in lower_list:
                    index_list.append(x[2])

                project_baseline_info[name] = list(lower_list)
                project_baseline_index[name] = list(index_list)

            baseline_info[baseline_dict[b_type]] = project_baseline_info
            baseline_index[baseline_dict[b_type]] = project_baseline_index

        self.bl_info = baseline_info
        self.bl_index = baseline_index

    def check_project_information(self) -> None:
        """Checks that project names in master are present/the same as in project info.
        Stops the programme if not"""
        error_cases = []
        for p in self.current_projects:
            if p not in self.all_projects:
                error_cases.append(p)

        if error_cases:
            for p in error_cases:
                logger.critical(p + " has not been found in the project_info document.")
            try:
                m = str(self.master_data[0].month)
            except KeyError:
                m = str(self.master_data[0].quarter)
            raise ProjectNameError(
                "Project names in the "
                + m
                + " master and project_info must match. Program stopping. Please amend."
            )
        else:
            logger.info("The latest master and project information match")

    def check_baselines(self) -> None:  # check with team is required for IPDC.
        """checks that projects have the correct baseline information. stops the
        programme if baselines are missing"""

        if "data_type" in self.kwargs:
            if self.kwargs["data_type"] == "cdg":
                baseline_dict = CDG_BASELINE_TYPES
        else:
            baseline_dict = IPDC_BASELINE_TYPES

        b_e_cases = []  # baseline error cases
        b_v_e_cases = []  # baseline value error cases
        for v in baseline_dict.values():
            for p in self.current_projects:
                baselines = self.bl_index[v][p]
                if len(baselines) <= 2:
                    b_e_cases.append(p)
                    if v not in b_v_e_cases:
                        b_v_e_cases.append(v)

        if b_e_cases:
            for i, b in enumerate(b_e_cases):
                logger.critical(
                    b_e_cases[i]
                    + " does not have a baseline point for "
                    + b_v_e_cases[i]
                    + " this could cause the programme to "
                      "crash. Therefore the programme is stopping. "
                      "Please amend the data for " + b_e_cases[i] + " so "
                        " it has at least one baseline point for " + b_v_e_cases[i]
                )
            raise ProjectNameError(  # should be Baselining Error or Initiation Error
                'Above issue(s) could cause a crash and require resolution. Program stopping'
            )

    ## Refactor required. Is this even used now?
    def get_project_groups(self) -> None:
        """gets the groups that projects are part of e.g. business case
        stage or dft group"""

        if "data_type" in self.kwargs:
            if self.kwargs["data_type"] == "cdg":
                group_key = "Directorate"
                # group_dict = CDG_DIR_DICT
                approval = "Last Business Case (BC) achieved"
            if self.kwargs["data_type"] == "top35":
                group_key = "Group"
                # group_dict = DFT_GROUP_DICT
            if self.kwargs["data_type"] == 'ipdc':
                group_key = "Group"
                # group_dict = DFT_GROUP_DICT
                approval = "IPDC approval point"

        raw_dict = {}
        raw_list = []
        group_list = []
        stage_list = []
        pn_e_cases = []  # project name error_cases
        p_m_e_cases = []  # project master error cases
        g_e_cases = []  # group error cases
        for i, master in enumerate(self.master_data):
            lower_dict = {}
            for p in master.projects:
                try:
                    dft_group = self.project_information[p][
                        group_key
                    ]
                except KeyError:
                    dft_group = None
                    pn_e_cases.append(p)
                    p_m_e_cases.append(str(master.quarter))

                if dft_group is None or dft_group not in self.all_groups:
                    g_e_cases.append(p)

                try:
                    stage = BC_STAGE_DICT[master[p][approval]]
                except (UnboundLocalError, NameError):  # top35 does not collect stage
                    stage = "None"
                raw_list.append(("group", dft_group))
                raw_list.append(("stage", stage))
                lower_dict[p] = dict(raw_list)
                group_list.append(dft_group)
                stage_list.append(stage)

            if pn_e_cases:
                for i, e in enumerate(pn_e_cases):
                    logger.critical(
                        f'Project name {pn_e_cases[i]} in master {p_m_e_cases[i]} not in project information '
                        f'document. Make sure project names are consistent.'
                    )
                raise ProjectNameError(
                    'Above issue(s) could cause a crash and require resolution. Program stopping'
                )

            if g_e_cases:
                for i in g_e_cases:
                    logger.critical(
                        str(i)
                        + " does not have a recognised Group value in the project information document."
                    )
                raise ProjectGroupError(
                    'Above issue(s) could cause a crash and require resolution. Program stopping'
                )


            try:
                raw_dict[str(master.month) + ", " + str(master.year)] = lower_dict
            except KeyError:
                raw_dict[str(master.quarter)] = lower_dict

        group_list = list(set(group_list))
        stage_list = list(set(stage_list))

        group_dict = {}
        # for i, quarter in enumerate(list(raw_dict.keys())[:2]):  # just latest two quarters
        for i, quarter in enumerate(list(raw_dict.keys())):
            lower_g_dict = {}
            for group_type in group_list:
                g_list = []
                for p in raw_dict[quarter].keys():
                    p_group = raw_dict[quarter][p]["group"]
                    if p_group == group_type:
                        g_list.append(p)
                lower_g_dict[group_type] = g_list

            gmpp_list = []
            for p in self.master_data[i].projects:
                try:
                    gmpp = self.project_information[p]["GMPP"]
                except KeyError:  # project name check happening in other places.
                    gmpp = None
                if gmpp is not None:
                    gmpp_list.append(p)
                lower_g_dict["GMPP"] = gmpp_list

            group_dict[quarter] = lower_g_dict

        stage_dict = {}
        for quarter in list(raw_dict.keys())[:2]:  # just latest two quarters
            lower_s_dict = {}
            for stage_type in stage_list:
                s_list = []
                for p in raw_dict[quarter].keys():
                    p_stage = raw_dict[quarter][p]["stage"]
                    if p_stage == stage_type:
                        s_list.append(p)
                if stage_type is None:
                    if s_list:
                        if "data_type" in self.kwargs:
                            if self.kwargs["data_type"] == "cdg":
                                continue  # not actively using stages for cdg data yet so can pass
                        if quarter == self.current_quarter:
                            for x in s_list:
                                logger.critical(str(x) + " has no IPDC stage date")
                                raise ProjectStageError(
                                    "Programme stopping as this could cause incomplete analysis"
                                )
                        else:
                            for x in s_list:
                                logger.warning(
                                    "In "
                                    + str(quarter)
                                    + " master "
                                    + str(x)
                                    + " IPDC stage data is currently None. Please amend."
                                )
                lower_s_dict[stage_type] = s_list
            stage_dict[quarter] = lower_s_dict

        self.dft_groups = group_dict
        self.project_stage = stage_dict

    def get_quarter_list(self) -> None:
        output_list = []
        for master in self.master_data:
            try:
                output_list.append(str(master.month) + ", " + str(master.year))
            except KeyError:
                output_list.append(str(master.quarter))
        self.quarter_list = output_list

    def pipeline_projects_information(self) -> None:
        pipeline_dict = {}
        pipeline_list = []
        total_wlc = 0
        for p in self.all_projects:
            if self.project_information[p]["Pipeline"] is not None:
                wlc = convert_none_types(self.project_information[p]["WLC"])
                pipeline_dict[p] = {
                    "wlc": convert_none_types(self.project_information[p]["WLC"])
                }
                pipeline_list.append(p)
                total_wlc += wlc
        pipeline_dict["pipeline"] = {"wlc": total_wlc}

        self.pipeline_dict = pipeline_dict
        self.pipeline_list = pipeline_list

    def get_current_tp(self):
        try:
            self.current_quarter = (
                    str(self.master_data[0].month) + ", " + str(self.master_data[0].year)
            )
        except KeyError:
            self.current_quarter = self.master_data[0].quarter


def json_date_converter(o):
    if isinstance(o, datetime.date):
        return o.__str__()


class JsonData:
    def __init__(
            self,
            master: List[Dict[str, Union[str, int, datetime.date, float]]],
            save_path: str
        ):
        self.master = master
        self.path = save_path
        self.put_into_json()

    def put_into_json(self) -> None:
        master_list = []
        for m in self.master.master_data:
            data = m.data
            projects = m.projects
            qrt = str(str(m.quarter))
            d = {
                "data": data,
                "projects": projects,
                "quarter": qrt,
            }
            master_list.append(d)

        json_dict = {
            "abbreviations": self.master.abbreviations,
            "bl_index": self.master.bl_index,
            "bl_info": self.master.bl_info,
            "current_projects": self.master.current_projects,
            "current_quarter": str(self.master.current_quarter),
            "dft_groups": self.master.dft_groups,
            "full_names": self.master.full_names,
            "kwargs": self.master.kwargs,
            "master_data": master_list,
            "pipeline_dict": self.master.pipeline_dict,
            "pipeline_list": self.master.pipeline_list,
            "project_group": self.master.project_group,
            "project_information": self.master.project_information,
            "project_stage": self.master.project_stage,
            "quarter_list": self.master.quarter_list,
        }
        with open(self.path + ".json", "w") as write_file:
            json.dump(json_dict, write_file, default=json_date_converter)


def get_group_stage_data(
    confi_path: Path,
) -> List[str]:
    # Returns a list of dft groups
    try:
        print(confi_path)
        config = configparser.ConfigParser()
        config.read(confi_path)
        # master_data_list = []
        portfolio_group = json.loads(
            config.get("GROUPS", "portfolio_groups")
        )  # to return a list
        group_all = json.loads(config.get("GROUPS", "all_groups"))
        try:
            bc_stages = json.loads(config.get("GROUPS", "bc_stages"))
        except configparser.NoOptionError:
            bc_stages = []
    except:
        logger.critical(
            "Configuration file issue. Please check and make sure it's correct."
        )
        sys.exit(1)

    return portfolio_group, group_all, bc_stages


def get_project_info_data(master_file: str) -> Dict:
    # taken from datamaps project_data_from_master
    wb = load_workbook(master_file)
    ws = wb.active
    for cell in ws["A"]:
        # we don't want to clean None...
        if cell.value is None:
            continue
        c = Cleanser(cell.value)
        cell.value = c.clean()
    p_dict = {}
    for col in ws.iter_cols(min_col=2):
        project_name = ""
        o = OrderedDict()
        for cell in col:
            if cell.row == 1:
                project_name = cell.value
                p_dict[project_name] = o
            else:
                val = ws.cell(row=cell.row, column=1).value
                if type(cell.value) == datetime:
                    d_value = datetime.date(cell.value.year, cell.value.month, cell.value.day)
                    p_dict[project_name][val] = d_value
                else:
                    p_dict[project_name][val] = cell.value
    # remove any "None" projects that were pulled from the master
    try:
        del p_dict[None]
    except KeyError:
        pass
    return p_dict


def get_project_information(
        confi_path: Path,
        pi_path: Path,
) -> Dict[str, Union[str, int]]:
    """Returns dictionary containing all project meta data.
    confi_path is ini file path.
    pi_path is project_info path."""
    config = configparser.ConfigParser()
    config.read(confi_path)
    path = str(pi_path) + config["PROJECT INFO"]["projects"]
    return get_project_info_data(path)


def get_core(
        reporting_type: str,
        config_file: str,
        func: Callable, #  project_data_from_master
    ) -> None:
    root_path = _platform_docs_dir(reporting_type)
    print(root_path)
    config_path = str(root_path) + config_file
    META = get_group_stage_data(config_path)
    all_groups = META[1]   # only all_groups used at initiate

    try:
        master = JsonMaster(
            get_master_data(
                config_path,
                str(root_path) + "/core_data/",
                func,
            ),
            get_project_information(
                config_path,
                str(root_path) + "/core_data/",
            ),
            all_groups,
            data_type=reporting_type
        )

    except (ProjectNameError, ProjectGroupError, ProjectStageError) as e:
        logger.critical(e)
        sys.exit(1)

    master_json_path = str("{0}/core_data/json/master".format(root_path))
    JsonData(master, master_json_path)