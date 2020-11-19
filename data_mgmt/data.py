import datetime
import difflib
import os
import re
import typing
from collections import Counter
from typing import List, Dict, Union, Optional, Tuple

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import timedelta, date
import numpy as np
from datamaps.api import project_data_from_master
import platform
from pathlib import Path

from datamaps.api.api import project_data_from_master_api
from docx import Document, table
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, Cm, RGBColor, Inches
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook import workbook


def _platform_docs_dir() -> Path:
    #  Cross plaform file path handling
    if platform.system() == "Linux":
        return Path.home() / "Documents" / "analysis_engine"
    if platform.system() == "Darwin":
        return Path.home() / "Documents" / "analysis_engine"
    else:
        return Path.home() / "Documents" / "analysis_engine"


root_path = _platform_docs_dir()


def get_master_data() -> List[
    Dict[str, Union[str, int, date, float]]
]:  # how specify a list of dictionaries?
    """Returns a list of dictionaries each containing quarter data"""
    master_data_list = [
        project_data_from_master(root_path / "core_data/master_2_2020.xlsx", 2, 2020),
        project_data_from_master(root_path / "core_data/master_1_2020.xlsx", 1, 2020),
        project_data_from_master(root_path / "core_data/master_4_2019.xlsx", 4, 2019),
        project_data_from_master(root_path / "core_data/master_3_2019.xlsx", 3, 2019),
        project_data_from_master(root_path / "core_data/master_2_2019.xlsx", 2, 2019),
        project_data_from_master(root_path / "core_data/master_1_2019.xlsx", 1, 2019),
        project_data_from_master(root_path / "core_data/master_4_2018.xlsx", 4, 2018),
        project_data_from_master(root_path / "core_data/master_3_2018.xlsx", 3, 2018),
        project_data_from_master(root_path / "core_data/master_2_2018.xlsx", 2, 2018),
        project_data_from_master(root_path / "core_data/master_1_2018.xlsx", 1, 2018),
        project_data_from_master(root_path / "core_data/master_4_2017.xlsx", 4, 2017),
        project_data_from_master(root_path / "core_data/master_3_2017.xlsx", 3, 2017),
        project_data_from_master(root_path / "core_data/master_2_2017.xlsx", 2, 2017),
        project_data_from_master(root_path / "core_data/master_1_2017.xlsx", 1, 2017),
        project_data_from_master(root_path / "core_data/master_4_2016.xlsx", 4, 2016),
        project_data_from_master(root_path / "core_data/master_3_2016.xlsx", 3, 2016),
    ]

    return master_data_list


def get_master_data_file_paths() -> List[typing.TextIO]:
    file_list = [
        root_path / "core_data/master_2_2020.xlsx",
        root_path / "core_data/master_1_2020.xlsx",
        root_path / "core_data/master_4_2019.xlsx",
        root_path / "core_data/master_3_2019.xlsx",
        root_path / "core_data/master_2_2019.xlsx",
        root_path / "core_data/master_1_2019.xlsx",
        root_path / "core_data/master_4_2018.xlsx",
        root_path / "core_data/master_3_2018.xlsx",
        root_path / "core_data/master_2_2018.xlsx",
        root_path / "core_data/master_1_2018.xlsx",
        root_path / "core_data/master_4_2017.xlsx",
        root_path / "core_data/master_3_2017.xlsx",
        root_path / "core_data/master_2_2017.xlsx",
        root_path / "core_data/master_1_2017.xlsx",
        root_path / "core_data/master_4_2016.xlsx",
        root_path / "core_data/master_3_2016.xlsx",
    ]

    return file_list


def get_master_data_file_paths_fy_16_17() -> List[typing.TextIO]:
    file_list = [
        root_path / "core_data/master_4_2016.xlsx",
        root_path / "core_data/master_3_2016.xlsx",
    ]
    return file_list


def get_master_data_file_paths_fy_17_18() -> List[typing.TextIO]:
    file_list = [
        root_path / "core_data/master_4_2017.xlsx",
        root_path / "core_data/master_3_2017.xlsx",
        root_path / "core_data/master_2_2017.xlsx",
        root_path / "core_data/master_1_2017.xlsx",
    ]
    return file_list


def get_master_data_file_paths_fy_18_19() -> List[typing.TextIO]:
    file_list = [
        root_path / "core_data/master_4_2018.xlsx",
        root_path / "core_data/master_3_2018.xlsx",
        root_path / "core_data/master_2_2018.xlsx",
        root_path / "core_data/master_1_2018.xlsx",
    ]
    return file_list


def get_master_data_file_paths_fy_19_20() -> List[typing.TextIO]:
    file_list = [
        root_path / "core_data/master_4_2019.xlsx",
        root_path / "core_data/master_3_2019.xlsx",
        root_path / "core_data/master_2_2019.xlsx",
        root_path / "core_data/master_1_2019.xlsx",
    ]

    return file_list


def get_master_data_file_paths_fy_20_21() -> List[typing.TextIO]:
    file_list = [
        root_path / "core_data/master_2_2020.xlsx",
        root_path / "core_data/master_1_2020.xlsx",
    ]

    return file_list


def get_datamap_file_paths():
    pass


def get_project_information() -> Dict[str, Union[str, int]]:
    """Returns dictionary containing all project meta data"""
    return project_data_from_master(
        root_path / "core_data/data_mgmt/project_info.xlsx", 2, 2020
    )


def get_project_information_file_path() -> typing.TextIO:
    return root_path / "core_data/other/project_info.xlsx"


def get_gmpp_projects(project_info_dict) -> List[str]:
    # returns list of projects that are in gmpp
    project_list = list(project_info_dict.data.keys())
    output_list = []
    for p in project_list:
        if project_info_dict.data[p]["GMPP"]:
            output_list.append(p)

    return output_list


# for project summary pages
SRO_CONF_TABLE_LIST = [
    "SRO DCA",
    "Finance DCA",
    "Benefits DCA",
    "Resourcing DCA",
    "Schedule DCA",
]
SRO_CONF_KEY_LIST = [
    "Departmental DCA",
    "SRO Finance confidence",
    "SRO Benefits RAG",
    "Overall Resource DCA - Now",
    "SRO Schedule Confidence",
]

IPDC_DATE = datetime.date(
    2020, 8, 10
)  # ipdc date. Python date format is Year, Month, day
blue_line_date = datetime.date.today()  # blue line on graph date.

# abbreviations. Used in analysis instead of full projects names
abbreviations = {
    "2nd Generation UK Search and Rescue Aviation": "SARH2",
    "A12 Chelmsford to A120 widening": "A12",
    "A14 Cambridge to Huntingdon Improvement Scheme": "A14",
    "A303 Amesbury to Berwick Down": "A303",
    "A358 Taunton to Southfields Dualling": "A358",
    "A417 Air Balloon": "A417",
    "A428 Black Cat to Caxton Gibbet": "A428",
    "A66 Northern Trans-Pennine": "A66",
    "Crossrail Programme": "Crossrail",
    "East Coast Digital Programme": "ECDP",
    "East Coast Mainline Programme": "ECMP",
    "East West Rail Programme (Central Section)": "EWR (Central)",
    "East West Rail Programme (Western Section)": "EWR (Western)",
    "East West Rail Configuration State 1": "EWR Config 1",
    "East West Rail Configuration State 2": "EWR Config 2",
    "East West Rail Configuration State 3": "EWR Config 3",
    "Future Theory Test Service (FTTS)": "FTTS",
    "Great Western Route Modernisation (GWRM) including electrification": "GWRM",
    "Heathrow Expansion": "HEP",
    "Hexagon": "Hexagon",
    "High Speed Rail Programme (HS2)": "HS2 Prog",
    "HS2 Phase 2b": "HS2 2b",
    "HS2 Phase1": "HS2 1",
    "HS2 Phase2a": "HS2 2a",
    "Integrated and Smart Ticketing - creating an account based back office": "IST",
    "Intercity Express Programme": "IEP",
    "Lower Thames Crossing": "LTC",
    "M4 Junctions 3 to 12 Smart Motorway": "M4",
    "Manchester North West Quadrant": "MNWQ",
    "Midland Main Line Programme": "MML Prog",
    "Midlands Rail Hub": "Mid Rail Hub",
    "North Western Electrification": "NWE",
    "Northern Powerhouse Rail": "NPR",
    "Oxford-Cambridge Expressway": "Ox-Cam Expressway",
    "Rail Franchising Programme": "Rail Franchising",
    "South West Route Capacity": "SWRC",
    "Thameslink Programme": "Thameslink",
    "Transpennine Route Upgrade (TRU)": "TRU",
    "Western Rail Link to Heathrow": "WRLtH",
}


class Projects:
    # project names as variables
    a12 = "A12 Chelmsford to A120 widening"
    a14 = "A14 Cambridge to Huntingdon Improvement Scheme"
    a303 = "A303 Amesbury to Berwick Down"
    a385 = "A358 Taunton to Southfields Dualling"
    a417 = "A417 Air Balloon"
    a428 = "A428 Black Cat to Caxton Gibbet"
    a66 = "A66 Northern Trans-Pennine"
    brighton_ml = "Brighton Mainline Upgrade Programme"
    cvs = "Commercial Vehicle Services (CVS)"
    east_coast_digital = "East Coast Digital Programme"
    east_coast_mainline = "East Coast Mainline Programme"
    em_franchise = "East Midlands Franchise"
    ewr_central = "East West Rail Programme (Central Section)"
    ewr_western = "East West Rail Programme (Western Section)"
    ewr_config1 = "East West Rail Configuration State 1"
    ewr_config2 = "East West Rail Configuration State 2"
    ewr_config3 = "East West Rail Configuration State 3"
    ftts = "Future Theory Test Service (FTTS)"
    heathrow_expansion = "Heathrow Expansion"
    hexagon = "Hexagon"
    hs2_programme = "High Speed Rail Programme (HS2)"
    hs2_2b = "HS2 Phase 2b"
    hs2_1 = "HS2 Phase1"
    hs2_2a = "HS2 Phase2a"
    ist = "Integrated and Smart Ticketing - creating an account based back office"
    lower_thames_crossing = "Lower Thames Crossing"
    m4 = "M4 Junctions 3 to 12 Smart Motorway"
    manchester_north_west_quad = "Manchester North West Quadrant"
    midland_mainline = "Midland Main Line Programme"
    midlands_rail_hub = "Midlands Rail Hub"
    north_of_england = "North of England Programme"
    northern_powerhouse = "Northern Powerhouse Rail"
    nwe = "North Western Electrification"
    ox_cam_expressway = "Oxford-Cambridge Expressway"
    rail_franchising = "Rail Franchising Programme"
    west_coast_partnership = "West Coast Partnership Franchise"
    crossrail = "Crossrail Programme"
    gwrm = "Great Western Route Modernisation (GWRM) including electrification"
    iep = "Intercity Express Programme"
    sarh2 = "2nd Generation UK Search and Rescue Aviation"
    south_west_route_capacity = "South West Route Capacity"
    thameslink = "Thameslink Programme"
    tru = "Transpennine Route Upgrade"
    wrlth = "Western Rail Link to Heathrow"

    # lists of projects names in groups
    he = [lower_thames_crossing, a303, a14, a66, a12, m4, a428, a417, a385]
    rail = [
        crossrail,
        thameslink,
        iep,
        east_coast_mainline,
        east_coast_digital,
        midland_mainline,
        nwe,
        south_west_route_capacity,
        brighton_ml,
        midlands_rail_hub,
        gwrm,
        tru,
        wrlth,
    ]
    hs2 = [hs2_1, hs2_2a, hs2_2b]
    hsmrpg = [
        hs2_1,
        hs2_2a,
        hs2_2b,
        ewr_config1,
        ewr_config2,
        ewr_config3,
        hexagon,
        northern_powerhouse,
    ]
    ewr = [ewr_config1, ewr_config2, ewr_config3]
    dvsa = [ftts, ist]
    all_not_hs2 = [
        "2nd Generation UK Search and Rescue Aviation",
        "A12 Chelmsford to A120 widening",
        "A14 Cambridge to Huntingdon Improvement Scheme",
        "A303 Amesbury to Berwick Down",
        "A358 Taunton to Southfields Dualling",
        "A417 Air Balloon",
        "A428 Black Cat to Caxton Gibbet",
        "A66 Northern Trans-Pennine",
        "Brighton Mainline Upgrade Programme",
        "Crossrail Programme",
        "East Coast Digital Programme",
        "East Coast Mainline Programme",
        "East West Rail Configuration State 1",
        "East West Rail Configuration State 2",
        "East West Rail Configuration State 3",
        "Future Theory Test Service (FTTS)",
        "Great Western Route Modernisation (GWRM) including electrification",
        "Hexagon",
        "Integrated and Smart Ticketing - creating an account based back office",
        "Intercity Express Programme",
        "Lower Thames Crossing",
        "M4 Junctions 3 to 12 Smart Motorway",
        "Midland Main Line Programme",
        "Midlands Rail Hub",
        "North Western Electrification",
        "Northern Powerhouse Rail",
        "Rail Franchising Programme",
        "South West Route Capacity",
        "Thameslink Programme",
        "Transpennine Route Upgrade",
        "Western Rail Link to Heathrow",
    ]
    fbc_stage = [
        hs2_1,
        crossrail,
        east_coast_mainline,
        iep,
        thameslink,
        south_west_route_capacity,
        hexagon,
        gwrm,
        nwe,
        midland_mainline,
        m4,
        a14,
        ewr_config1,
    ]
    obc_stage = [
        lower_thames_crossing,
        hs2_2a,
        tru,
        east_coast_digital,
        a303,
        a12,
        a428,
        a417,
        a385,
        ftts,
    ]
    sobc_stage = [
        hs2_2b,
        brighton_ml,
        ewr_config3,
        sarh2,
        midlands_rail_hub,
        wrlth,
        a66,
        ewr_config2,
    ]


#  list of different baseline types. hold at global level?
BASELINE_TYPES = {
    "Re-baseline this quarter": "quarter",
    "Re-baseline ALB/Programme milestones": "programme_milestones",
    "Re-baseline ALB/Programme cost": "programme_costs",
    "Re-baseline ALB/Programme benefits": "programme_benefits",
    "Re-baseline IPDC milestones": "ipdc_milestones",
    "Re-baseline IPDC cost": "ipdc_costs",
    "Re-baseline IPDC benefits": "ipdc_benefits",
    "Re-baseline HMT milestones": "hmt_milestones",
    "Re-baseline HMT cost": "hmt_costs",
    "Re-baseline HMT benefits": "hmt_benefits",
}

YEAR_LIST = [
    "16-17",
    "17-18",
    "18-19",
    "19-20",
    "20-21",
    "21-22",
    "22-23",
    "23-24",
    "24-25",
    "25-26",
    "26-27",
    "27-28",
    "28-29",
    "29-30",
    "30-31",
    "31-32",
    "32-33",
    "33-34",
    "34-35",
    "35-36",
    "36-37",
    "37-38",
    "38-39",
    "39-40",
]

COST_KEY_LIST = [
    " RDEL Forecast Total",
    " CDEL Forecast one off new costs",
    " Forecast Non-Gov",
]
COST_TYPE_KEY_LIST = [
    (
        "Pre-profile RDEL",
        "Pre-profile CDEL Forecast one off new costs",
        "Pre-profile Forecast Non-Gov",
    ),
    (
        "Total RDEL Forecast Total",
        "Total CDEL Forecast one off new costs",
        "Non-Gov Total Forecast",
    ),
    (
        "Unprofiled RDEL Forecast Total",
        "Unprofiled CDEL Forecast one off new costs",
        "Unprofiled Forecast Non-Gov",
    ),
]
BEN_KEY_LIST = [
    "Pre-profile BEN Total",
    "Total BEN Forecast - Total Monetised Benefits",
    "Unprofiled Remainder BEN Forecast - Total Monetised Benefits",
]
BEN_TYPE_KEY_LIST = [
    (
        "Pre-profile BEN Forecast Gov Cashable",
        "Pre-profile BEN Forecast Gov Non-Cashable",
        "Pre-profile BEN Forecast - Economic (inc Private Partner)",
        "Pre-profile BEN Forecast - Disbenefit UK Economic",
    ),
    (
        "Total BEN Forecast - Gov. Cashable",
        "Total BEN Forecast - Gov. Non-Cashable",
        "Total BEN Forecast - Economic (inc Private Partner)",
        "Total BEN Forecast - Disbenefit UK Economic",
    ),
    (
        "Unprofiled Remainder BEN Forecast - Gov. Cashable",
        "Unprofiled Remainder BEN Forecast - Gov. Non-Cashable",
        "Unprofiled Remainder BEN Forecast - Economic (inc Private Partner)",
        "Unprofiled Remainder BEN Forecast - Disbenefit UK Economic",
    ),
]

# Matplotlib file formats
FILE_FORMATS = [
    "eps",
    "jpeg",
    "jpg",
    "pdf",
    "png",
    "ps",
    "raw",
    "rgba",
    "svg",
    "svgz",
    "tif",
    "tiff",
]

FIGURE_STYLE = {1: "half horizontal", 2: "full horizontal"}


def calculate_profiled(p: List[int], s: List[int], unpro: List[int]) -> list:
    """small helper function to calculate the proper profiled amount. This is necessary as
    other wise 'profiled' would actually be the total figure.
    p = profiled list
    s = spent list
    unpro = unprofiled list"""
    f_profiled = []
    for y, amount in enumerate(p):
        t = amount - (s[y] + unpro[y])
        f_profiled.append(t)
    return f_profiled


class Master:
    def __init__(
        self,
        master_data: List[Dict[str, Union[str, int, date, float]]],
        project_information: Dict[str, Union[str, int]],
    ) -> None:
        self.master_data = master_data
        self.project_information = project_information
        self.current_projects = self.get_current_projects()
        self.check_project_information()
        self.bl_info = {}
        self.bl_index = {}
        self.baseline_data()
        # self.check_baselines()  # optional for now

    def baseline_data(self) -> dict:
        """
        Returns the two dictionaries baseline_info and baseline_index for all projects for all
        baseline types
        """

        baseline_info = {}
        baseline_index = {}

        for b_type in list(BASELINE_TYPES.keys()):
            project_baseline_info = {}
            project_baseline_index = {}
            for name in self.current_projects:
                bc_list = []
                lower_list = []
                for i, master in reversed(list(enumerate(self.master_data))):
                    if name in master.projects:
                        try:
                            approved_bc = master.data[name][b_type]
                            quarter = str(master.quarter)
                        except KeyError:  # exception handling in here because data keys across masters are not consistent.
                            print(
                                str(b_type)
                                + " key not present in "
                                + str(master.quarter)
                            )
                        if approved_bc == "Yes":
                            bc_list.append(approved_bc)
                            lower_list.append((approved_bc, quarter, i))
                    else:
                        pass
                for i in reversed(range(2)):
                    if name in self.master_data[i].projects:
                        approved_bc = self.master_data[i][name][b_type]
                        quarter = str(self.master_data[i].quarter)
                        lower_list.append((approved_bc, quarter, i))
                    else:
                        quarter = str(self.master_data[i].quarter)
                        lower_list.append((None, quarter, None))

                index_list = []
                for x in lower_list:
                    index_list.append(x[2])

                project_baseline_info[name] = list(reversed(lower_list))
                project_baseline_index[name] = list(reversed(index_list))

            baseline_info[BASELINE_TYPES[b_type]] = project_baseline_info
            baseline_index[BASELINE_TYPES[b_type]] = project_baseline_index

        self.bl_info = baseline_info
        self.bl_index = baseline_index

    def get_current_projects(self) -> list:
        """Returns a list of all the project names in the latest master"""
        return self.master_data[0].projects

    def check_project_information(self) -> None:
        """Checks that project names in master are present/the same as in project info.
        Stops the programme if not"""
        for p in self.current_projects:
            if p not in self.project_information.projects:
                print(
                    p
                    + " is not in the projects information document. Project names must be identical "
                    " in both documents. Programme stopping. Please amend."
                )
                break
            else:
                if p == self.current_projects[-1]:
                    print("The latest master and project information match")
                continue

    def check_baselines(self) -> None:
        """checks that projects have the correct baseline information. stops the
        programme if baselines are missing"""
        # work through best way to stop the programme.
        for v in BASELINE_TYPES.values():
            for p in self.current_projects:
                baselines = self.bl_index[v][p]
                if len(baselines) <= 2:
                    print(
                        p
                        + " does not have a baseline point for "
                        + v
                        + " this could cause the programme to "
                        "crash. Therefore the programme is stopping. "
                        "Please amend the data for " + p + " so that "
                        " it has at least one baseline point for " + v
                    )
            else:
                continue
            break


class MilestoneData:
    def __init__(self, masters_object, abbreviations):
        self.masters = masters_object
        self.abbreviations = abbreviations
        self.project_current = {}
        self.project_last = {}
        self.project_baseline = {}
        self.project_baseline_two = {}
        self.group_current = {}
        self.group_last = {}
        self.group_baseline = {}
        self.group_baseline_two = {}
        self.project_choronological_list_current = []
        self.project_choronological_list_last = []
        self.project_choronological_list_baseline = []
        self.group_choronological_list_current = []
        self.group_choronological_list_last = []
        self.group_choronological_list_baseline = []
        # self.project_data()
        # self.group_data()

    def project_data(self, milestone_type):  # renamed to figproject_data
        """
        Creates project milestone dictionaries for current, last, and
        baselines when provided with a milestone_type for all
        projects within a MilestoneData type.
        """
        self.milestone_type = milestone_type

        current_dict = {}
        last_dict = {}
        baseline_dict = {}
        baseline_dict_two = {}
        sorted_current = []
        sorted_last = []
        sorted_baseline = []

        if self.milestone_type == "All":
            for name in self.masters.current_projects:
                for ind in self.masters.bl_index[name][:4]:  # limit to four for now
                    lower_dict = {}
                    raw_list = []
                    try:
                        p_data = self.masters.master_data[ind].data[name]
                        for i in range(1, 50):
                            try:
                                try:
                                    t = (
                                        p_data["Approval MM" + str(i)],
                                        p_data[
                                            "Approval MM"
                                            + str(i)
                                            + " Forecast / Actual"
                                        ],
                                        p_data["Approval MM" + str(i) + " Notes"],
                                    )
                                    raw_list.append(t)
                                except KeyError:
                                    t = (
                                        p_data["Approval MM" + str(i)],
                                        p_data[
                                            "Approval MM"
                                            + str(i)
                                            + " Forecast - Actual"
                                        ],
                                        p_data["Approval MM" + str(i) + " Notes"],
                                    )
                                    raw_list.append(t)

                                t = (
                                    p_data["Assurance MM" + str(i)],
                                    p_data[
                                        "Assurance MM" + str(i) + " Forecast - Actual"
                                    ],
                                    p_data["Assurance MM" + str(i) + " Notes"],
                                )
                                raw_list.append(t)

                            except KeyError:
                                pass

                        for i in range(18, 67):
                            try:
                                t = (
                                    p_data["Project MM" + str(i)],
                                    p_data[
                                        "Project MM" + str(i) + " Forecast - Actual"
                                    ],
                                    p_data["Project MM" + str(i) + " Notes"],
                                )
                                raw_list.append(t)
                            except KeyError:
                                pass
                    except (KeyError, TypeError):
                        pass

                    # put the list in chronological order
                    sorted_list = sorted(raw_list, key=lambda k: (k[1] is None, k[1]))

                    # loop to stop key names being the same. Not ideal as doesn't handle keys that may already have numbers as
                    # strings at end of names. But still useful.
                    for x in sorted_list:
                        if x[0] is not None:
                            if x[0] in lower_dict:
                                for y in range(2, 15):
                                    key_name = x[0] + " " + str(y)
                                    if key_name in lower_dict:
                                        continue
                                    else:
                                        lower_dict[key_name] = {x[1]: x[2]}
                                        break
                            else:
                                lower_dict[x[0]] = {x[1]: x[2]}
                        else:
                            pass

                    if self.masters.bl_index[name].index(ind) == 0:
                        current_dict[name] = lower_dict
                        sorted_current = sorted_list
                    if self.masters.bl_index[name].index(ind) == 1:
                        last_dict[name] = lower_dict
                        sorted_last = sorted_list
                    if self.masters.bl_index[name].index(ind) == 2:
                        baseline_dict[name] = lower_dict
                        sorted_baseline = sorted_list
                    if self.masters.bl_index[name].index(ind) == 3:
                        baseline_dict_two[name] = lower_dict

        if self.milestone_type == "Delivery":
            for name in self.masters.current_projects:
                for ind in self.masters.bl_index[name][:4]:  # limit to four for now
                    lower_dict = {}
                    raw_list = []
                    try:
                        p_data = self.masters.master_data[ind].data[name]
                        for i in range(18, 67):
                            try:
                                t = (
                                    p_data["Project MM" + str(i)],
                                    p_data[
                                        "Project MM" + str(i) + " Forecast - Actual"
                                    ],
                                    p_data["Project MM" + str(i) + " Notes"],
                                )
                                raw_list.append(t)
                            except KeyError:
                                pass
                    except (KeyError, TypeError):
                        pass

                    # put the list in chronological order
                    sorted_list = sorted(raw_list, key=lambda k: (k[1] is None, k[1]))

                    # loop to stop key names being the same. Not ideal as doesn't handle keys that may already have numbers as
                    # strings at end of names. But still useful.
                    for x in sorted_list:
                        if x[0] is not None:
                            if x[0] in lower_dict:
                                for y in range(2, 15):
                                    key_name = x[0] + " " + str(y)
                                    if key_name in lower_dict:
                                        continue
                                    else:
                                        lower_dict[key_name] = {x[1]: x[2]}
                                        break
                            else:
                                lower_dict[x[0]] = {x[1]: x[2]}
                        else:
                            pass

                    if self.masters.bl_index[name].index(ind) == 0:
                        current_dict[name] = lower_dict
                        # sorted_current = sorted_list
                    if self.masters.bl_index[name].index(ind) == 1:
                        last_dict[name] = lower_dict
                        # sorted_last = sorted_list
                    if self.masters.bl_index[name].index(ind) == 2:
                        baseline_dict[name] = lower_dict
                        # sorted_baseline = sorted_list
                    if self.masters.bl_index[name].index(ind) == 3:
                        baseline_dict_two[name] = lower_dict

        self.project_current = current_dict
        self.project_last = last_dict
        self.project_baseline = baseline_dict
        self.project_baseline_two = baseline_dict_two
        self.project_choronological_list_current = sorted_current
        self.project_choronological_list_last = sorted_last
        self.project_choronological_list_baseline = sorted_baseline

    def group_data(self, milestone_type):
        """
        Creates group milestone dictionaries for current, last, and
        baselines when provided with a milestone_type for all
        projects within a MilestoneData type.
        """
        self.milestone_type = milestone_type

        current_dict = {}
        last_dict = {}
        baseline_dict = {}
        baseline_dict_two = {}
        sorted_current = []
        sorted_last = []
        sorted_baseline = []

        if self.milestone_type == "All":
            for num in range(0, 4):
                raw_list = []
                for name in self.masters.current_projects:
                    try:
                        p_data = self.masters.master_data[
                            self.masters.bl_index[name][num]
                        ].data[name]
                        for i in range(1, 50):
                            try:
                                try:
                                    if p_data["Approval MM" + str(i)] is None:
                                        pass
                                    else:
                                        key_name = (
                                            self.abbreviations[name]
                                            + ", "
                                            + p_data["Approval MM" + str(i)]
                                        )
                                        t = (
                                            key_name,
                                            p_data[
                                                "Approval MM"
                                                + str(i)
                                                + " Forecast / Actual"
                                            ],
                                            p_data["Approval MM" + str(i) + " Notes"],
                                        )
                                        raw_list.append(t)
                                except KeyError:
                                    if p_data["Approval MM" + str(i)] is None:
                                        pass
                                    else:
                                        key_name = (
                                            self.abbreviations[name]
                                            + ", "
                                            + p_data["Approval MM" + str(i)]
                                        )
                                        t = (
                                            key_name,
                                            p_data[
                                                "Approval MM"
                                                + str(i)
                                                + " Forecast - Actual"
                                            ],
                                            p_data["Approval MM" + str(i) + " Notes"],
                                        )
                                        raw_list.append(t)

                                if p_data["Assurance MM" + str(i)] is None:
                                    pass
                                else:
                                    key_name = (
                                        self.abbreviations[name]
                                        + ", "
                                        + p_data["Assurance MM" + str(i)]
                                    )
                                    t = (
                                        key_name,
                                        p_data[
                                            "Assurance MM"
                                            + str(i)
                                            + " Forecast - Actual"
                                        ],
                                        p_data["Assurance MM" + str(i) + " Notes"],
                                    )
                                    raw_list.append(t)

                            except KeyError:
                                pass

                        for i in range(18, 67):
                            try:
                                if p_data["Project MM" + str(i)] is None:
                                    pass
                                else:
                                    key_name = (
                                        self.abbreviations[name]
                                        + ", "
                                        + p_data["Project MM" + str(i)]
                                    )
                                    t = (
                                        key_name,
                                        p_data[
                                            "Project MM" + str(i) + " Forecast - Actual"
                                        ],
                                        p_data["Project MM" + str(i) + " Notes"],
                                    )
                                    raw_list.append(t)
                            except KeyError:
                                pass
                    except (KeyError, TypeError, IndexError):
                        pass

                sorted_list = sorted(
                    raw_list, key=lambda k: (k[1] is None, k[1])
                )  # put the list in chronological order

                """loop to stop key names being the same. Not ideal as doesn't handle keys that may
                already have numbers as strings at end of names. But still useful."""

                output_dict = {}
                for x in sorted_list:
                    if x[0] is not None:
                        if x[0] in output_dict:
                            for i in range(2, 15):
                                key_name = x[0] + " " + str(i)
                                if key_name in output_dict:
                                    continue
                                else:
                                    output_dict[key_name] = {x[1]: x[2]}
                                    break
                        else:
                            output_dict[x[0]] = {x[1]: x[2]}
                    else:
                        pass

                if num == 0:
                    current_dict = output_dict
                    sorted_current = sorted_list
                if num == 1:
                    last_dict = output_dict
                    sorted_last = sorted_list
                if num == 2:
                    baseline_dict = output_dict
                    sorted_baseline = sorted_list
                if num == 3:
                    baseline_dict_two = output_dict

        self.group_current = current_dict
        self.group_last = last_dict
        self.group_baseline = baseline_dict
        self.group_baseline_two = baseline_dict_two
        self.group_choronological_list_current = sorted_current
        self.group_choronological_list_last = sorted_last
        self.group_choronological_list_baseline = sorted_baseline


class MilestoneChartData:
    def __init__(
        self,
        milestone_data_object,
        keys_of_interest=None,
        keys_not_of_interest=None,
        filter_start_date=datetime.date(2000, 1, 1),
        filter_end_date=datetime.date(2050, 1, 1),
    ):
        self.m_data = milestone_data_object
        self.keys_of_interest = keys_of_interest
        self.keys_not_of_interest = keys_not_of_interest
        self.filter_start_date = filter_start_date
        self.filter_end_date = filter_end_date
        self.group_keys = []
        self.group_current_tds = []
        self.group_last_tds = []
        self.group_baseline_tds = []
        self.group_baseline_tds_two = []
        self.group_chart()

    def group_chart(self):
        """
        Given optional requirements, returns lists containing
        data for a group of project.
        key_of_interest is either default none or a list of strings
        """

        key_names = []
        td_current = []
        td_last = []
        td_baseline = []
        td_baseline_two = []

        # all milestone keys and time deltas calculated this way so
        # shown in particular way in output chart
        for m in list(self.m_data.group_current.keys()):
            if (
                "Project - Business Case End Date" in m
            ):  # filter out as dates not helpful
                pass
            else:
                if m is not None:
                    m_d_current = tuple(self.m_data.group_current[m])[0]

                if m in list(self.m_data.group_last.keys()):
                    m_d_last = tuple(self.m_data.group_last[m])[0]
                    if m_d_last is None:
                        m_d_last = tuple(self.m_data.group_current[m])[0]
                else:
                    m_d_last = tuple(self.m_data.group_current[m])[0]

                if m in list(self.m_data.group_baseline.keys()):
                    m_d_baseline = tuple(self.m_data.group_baseline[m])[0]
                    if m_d_baseline is None:
                        m_d_baseline = tuple(self.m_data.group_current[m])[0]
                else:
                    m_d_baseline = tuple(self.m_data.group_current[m])[0]

                if m in list(self.m_data.group_baseline_two.keys()):
                    m_d_baseline_two = tuple(self.m_data.group_baseline_two[m])[0]
                    if m_d_baseline_two is None:
                        m_d_baseline_two = tuple(self.m_data.group_current[m])[0]
                else:
                    m_d_baseline_two = tuple(self.m_data.group_current[m])[0]

                if m_d_current is not None:
                    if self.filter_start_date <= m_d_current <= self.filter_end_date:
                        if self.keys_of_interest is None:
                            key_names.append(m)
                            td_current.append(m_d_current)
                            td_last.append(m_d_last)
                            td_baseline.append(m_d_baseline)
                            td_baseline_two.append(m_d_baseline_two)

                        else:
                            for key in self.keys_of_interest:
                                if key in m:
                                    if m not in key_names:  # prevent repeats
                                        key_names.append(m)
                                        td_current.append(m_d_current)
                                        td_last.append(m_d_last)
                                        td_baseline.append(m_d_baseline)
                                        td_baseline_two.append(m_d_baseline_two)

        # loop to remove
        if self.keys_not_of_interest is not None:
            for x in range(len(key_names)):
                for y in self.keys_not_of_interest:
                    try:
                        if y in key_names[x]:
                            key_names[x] = None
                            td_current[x] = None
                            td_last[x] = None
                            td_baseline[x] = None
                            td_baseline_two[x] = None
                    except TypeError:
                        pass

        key_names_final = [x for x in key_names if x is not None]
        td_current_final = [x for x in td_current if x is not None]
        td_last_final = [x for x in td_last if x is not None]
        td_baseline_final = [x for x in td_baseline if x is not None]
        td_baseline_two_final = [x for x in td_baseline_two if x is not None]

        self.group_keys = key_names_final
        self.group_current_tds = td_current_final
        self.group_last_tds = td_last_final
        self.group_baseline_tds = td_baseline_final
        self.group_baseline_tds_two = td_baseline_two_final


class CombinedData:
    def __init__(self, wb, pfm_milestone_data):
        self.wb = wb
        self.pfm_milestone_data = pfm_milestone_data
        # self.project_current = {}
        # self.project_last = {}
        # self.project_baseline = {}
        # self.project_baseline_two = {}
        self.group_current = {}
        self.group_last = {}
        self.group_baseline = {}
        self.group_baseline_two = {}
        self.combined_tuple_list_forecast = []
        self.combined_tuple_list_baseline = []
        self.combine_mi_pfm_data()

    def combine_mi_pfm_data(self):
        """
        coverts data from MI system into usable format for graphical outputs
        """
        ws = self.wb.active

        mi_milestone_name_list = []  # handles duplicates
        mi_tuple_list_forecast = []
        mi_tuple_list_baseline = []
        for r in range(4, ws.max_row + 1):
            mi_milestone_key_name_raw = ws.cell(row=r, column=3).value
            mi_milestone_key_name = "MI, " + mi_milestone_key_name_raw
            forecast_date = ws.cell(row=r, column=8).value
            baseline_date = ws.cell(row=r, column=9).value
            notes = ws.cell(row=r, column=10).value
            if mi_milestone_key_name not in mi_milestone_name_list:
                mi_milestone_name_list.append(mi_milestone_key_name)
                mi_tuple_list_forecast.append(
                    (mi_milestone_key_name, forecast_date.date(), notes)
                )
                mi_tuple_list_baseline.append(
                    (mi_milestone_key_name, baseline_date.date(), notes)
                )
            else:
                for i in range(
                    2, 15
                ):  # alters duplicates by adding number to end of key
                    mi_altered_milestone_key_name = mi_milestone_key_name + " " + str(i)
                    if mi_altered_milestone_key_name in mi_milestone_name_list:
                        continue
                    else:
                        mi_tuple_list_forecast.append(
                            (mi_altered_milestone_key_name, forecast_date.date(), notes)
                        )
                        mi_tuple_list_baseline.append(
                            (mi_altered_milestone_key_name, baseline_date.date(), notes)
                        )
                        mi_milestone_name_list.append(mi_altered_milestone_key_name)
                        break

        mi_tuple_list_forecast = sorted(
            mi_tuple_list_forecast, key=lambda k: (k[1] is None, k[1])
        )  # put the list in chronological order
        mi_tuple_list_baseline = sorted(
            mi_tuple_list_baseline, key=lambda k: (k[1] is None, k[1])
        )  # put the list in chronological order

        pfm_tuple_list_forecast = []
        pfm_tuple_list_baseline = []
        for data in self.pfm_milestone_data.group_choronological_list_current:
            pfm_tuple_list_forecast.append(("PfM, " + data[0], data[1], data[2]))
        for data in self.pfm_milestone_data.group_choronological_list_baseline:
            pfm_tuple_list_baseline.append(("PfM, " + data[0], data[1], data[2]))

        combined_tuple_list_forecast = mi_tuple_list_forecast + pfm_tuple_list_forecast
        combined_tuple_list_baseline = mi_tuple_list_baseline + pfm_tuple_list_baseline

        combined_tuple_list_forecast = sorted(
            combined_tuple_list_forecast, key=lambda k: (k[1] is None, k[1])
        )  # put the list in chronological order
        combined_tuple_list_baseline = sorted(
            combined_tuple_list_baseline, key=lambda k: (k[1] is None, k[1])
        )  # put the list in chronological order

        milestone_dict_forecast = {}
        for x in combined_tuple_list_forecast:
            if x[0] is not None:
                milestone_dict_forecast[x[0]] = {x[1]: x[2]}
        milestone_dict_baseline = {}
        for x in combined_tuple_list_baseline:
            if x[0] is not None:
                milestone_dict_baseline[x[0]] = {x[1]: x[2]}

        self.group_current = milestone_dict_forecast
        self.group_last = {}
        self.group_baseline = milestone_dict_baseline
        self.group_baseline_two = {}
        self.combined_tuple_list_forecast = combined_tuple_list_forecast
        self.combined_tuple_list_baseline = combined_tuple_list_baseline


class MilestoneCharts:
    def __init__(
        self,
        latest_milestone_names,
        latest_milestone_dates,
        last_milestone_dates,
        baseline_milestone_dates,
        graph_title,
        ipdc_date,
    ):
        self.latest_milestone_names = latest_milestone_names
        self.latest_milestone_dates = latest_milestone_dates
        self.last_milestone_dates = last_milestone_dates
        self.baseline_milestone_dates = baseline_milestone_dates
        self.graph_title = graph_title
        self.ipdc_date = ipdc_date
        # self.milestone_swimlane_charts()
        self.build_charts()

    def milestone_swimlane_charts(self):
        # build scatter chart
        fig, ax1 = plt.subplots()
        fig.suptitle(self.graph_title, fontweight="bold")  # title
        # set fig size
        fig.set_figheight(4)
        fig.set_figwidth(8)

        ax1.scatter(
            self.baseline_milestone_dates, self.latest_milestone_names, label="Baseline"
        )
        ax1.scatter(
            self.last_milestone_dates, self.latest_milestone_names, label="Last Qrt"
        )
        ax1.scatter(
            self.latest_milestone_dates, self.latest_milestone_names, label="Latest Qrt"
        )

        # format the x ticks
        years = mdates.YearLocator()  # every year
        months = mdates.MonthLocator()  # every month
        years_fmt = mdates.DateFormatter("%Y")
        months_fmt = mdates.DateFormatter("%b")

        # calculate the length of the time period covered in chart. Not perfect as baseline dates can distort.
        try:
            td = (self.latest_milestone_dates[-1] - self.latest_milestone_dates[0]).days
            if td <= 365 * 3:
                ax1.xaxis.set_major_locator(years)
                ax1.xaxis.set_minor_locator(months)
                ax1.xaxis.set_major_formatter(years_fmt)
                ax1.xaxis.set_minor_formatter(months_fmt)
                plt.setp(ax1.xaxis.get_minorticklabels(), rotation=45)
                plt.setp(
                    ax1.xaxis.get_majorticklabels(), rotation=45, weight="bold"
                )  # milestone_swimlane_charts(key_name,
                #                           current_m_data,
                #                           last_m_data,
                #                           baseline_m_data,
                #                           'All Milestones')
                # scaling x axis
                # x axis value to no more than three months after last latest milestone date, or three months
                # before first latest milestone date. Hack, can be improved. Text highlights movements off chart.
                x_max = self.latest_milestone_dates[-1] + timedelta(days=90)
                x_min = self.latest_milestone_dates[0] - timedelta(days=90)
                for date in self.baseline_milestone_dates:
                    if date > x_max:
                        ax1.set_xlim(x_min, x_max)
                        plt.figtext(
                            0.98,
                            0.03,
                            "Check full schedule to see all milestone movements",
                            horizontalalignment="right",
                            fontsize=6,
                            fontweight="bold",
                        )
                    if date < x_min:
                        ax1.set_xlim(x_min, x_max)
                        plt.figtext(
                            0.98,
                            0.03,
                            "Check full schedule to see all milestone movements",
                            horizontalalignment="right",
                            fontsize=6,
                            fontweight="bold",
                        )
            else:
                ax1.xaxis.set_major_locator(years)
                ax1.xaxis.set_minor_locator(months)
                ax1.xaxis.set_major_formatter(years_fmt)
                plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, weight="bold")
        except IndexError:  # if milestone dates list is empty:
            pass

        ax1.legend()  # insert legend

        # reverse y axis so order is earliest to oldest
        ax1 = plt.gca()
        ax1.set_ylim(ax1.get_ylim()[::-1])
        ax1.tick_params(axis="y", which="major", labelsize=7)
        ax1.yaxis.grid()  # horizontal lines
        ax1.set_axisbelow(True)
        # ax1.get_yaxis().set_visible(False)

        # for i, txt in enumerate(latest_milestone_names):
        #     ax1.annotate(txt, (i, latest_milestone_dates[i]))

        # Add line of IPDC date, but only if in the time period
        try:
            if (
                self.latest_milestone_dates[0]
                <= self.ipdc_date
                <= self.latest_milestone_dates[-1]
            ):
                plt.axvline(self.ipdc_date)
                plt.figtext(
                    0.98,
                    0.01,
                    "Line represents when IPDC will discuss Q1 20_21 portfolio management report",
                    horizontalalignment="right",
                    fontsize=6,
                    fontweight="bold",
                )
        except IndexError:
            pass

        # size of chart and fit
        fig.canvas.draw()
        fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

        fig.savefig(
            root_path / "output/{}.png".format(self.graph_title), bbox_inches="tight"
        )

        # plt.close() #automatically closes figure so don't need to do manually.

    def build_charts(self):

        # add \n to y axis labels and cut down if two long
        # labels = ['\n'.join(wrap(l, 40)) for l in latest_milestone_names]
        labels = self.latest_milestone_names
        final_labels = []
        for l in labels:
            if len(l) > 40:
                final_labels.append(l[:35])
            else:
                final_labels.append(l)

        # Chart
        no_milestones = len(self.latest_milestone_names)

        if no_milestones <= 30:
            (
                np.array(final_labels),
                np.array(self.latest_milestone_dates),
                np.array(self.last_milestone_dates),
                np.array(self.baseline_milestone_dates),
                self.graph_title,
                self.ipdc_date,
            )

        if 31 <= no_milestones <= 60:
            half = int(no_milestones / 2)
            MilestoneCharts(
                np.array(final_labels[:half]),
                np.array(self.latest_milestone_dates[:half]),
                np.array(self.last_milestone_dates[:half]),
                np.array(self.baseline_milestone_dates[:half]),
                self.graph_title,
                self.ipdc_date,
            )
            title = self.graph_title + " cont."
            MilestoneCharts(
                np.array(final_labels[half:no_milestones]),
                np.array(self.latest_milestone_dates[half:no_milestones]),
                np.array(self.last_milestone_dates[half:no_milestones]),
                np.array(self.baseline_milestone_dates[half:no_milestones]),
                title,
                self.ipdc_date,
            )

        if 61 <= no_milestones <= 90:
            third = int(no_milestones / 3)
            MilestoneCharts(
                np.array(final_labels[:third]),
                np.array(self.latest_milestone_dates[:third]),
                np.array(self.last_milestone_dates[:third]),
                np.array(self.baseline_milestone_dates[:third]),
                self.graph_title,
                self.ipdc_date,
            )
            title = self.graph_title + " cont. 1"
            MilestoneCharts(
                np.array(final_labels[third : third * 2]),
                np.array(self.latest_milestone_dates[third : third * 2]),
                np.array(self.last_milestone_dates[third : third * 2]),
                np.array(self.baseline_milestone_dates[third : third * 2]),
                title,
                self.ipdc_date,
            )
            title = self.graph_title + " cont. 2"
            MilestoneCharts(
                np.array(final_labels[third * 2 : no_milestones]),
                np.array(self.latest_milestone_dates[third * 2 : no_milestones]),
                np.array(self.last_milestone_dates[third * 2 : no_milestones]),
                np.array(self.baseline_milestone_dates[third * 2 : no_milestones]),
                title,
                self.ipdc_date,
            )
        pass


class CostData:
    def __init__(self, master: Master):
        self.master = master
        self.cat_spent = []
        self.cat_profiled = []
        self.cat_unprofiled = []
        self.spent = []
        self.profiled = []
        self.unprofiled = []
        self.current_profile = []
        self.last_profile = []
        self.baseline_profile_one = []
        self.baseline_profile_two = []
        self.baseline_profile_three = []
        self.rdel_profile = []
        self.cdel_profile = []
        self.ngov_profile = []
        self.y_scale_max = 0
        self.entity = []

    def get_cost_totals(self, group: List[str] or str, baseline: str) -> None:
        """Returns lists containing the sum total of group (of projects) costs,
        sliced in different ways. Cumbersome for loop used at the moment, but
        is the least cumbersome loop I could design!"""
        self.group = group
        self.baseline = baseline

        spent = []
        profiled = []
        unprofiled = []
        group_rdel_spent = 0
        group_cdel_spent = 0
        group_ngov_spent = 0
        group_rdel_profiled = 0
        group_cdel_profiled = 0
        group_ngov_profiled = 0
        group_rdel_unprofiled = 0
        group_cdel_unprofiled = 0
        group_ngov_unprofiled = 0

        group = string_conversion(self.group)

        for i in range(3):
            for x, key in enumerate(COST_TYPE_KEY_LIST):
                group_total = 0
                for project_name in group:
                    cost_bl_index = self.master.bl_index[baseline][project_name]
                    try:
                        rdel = self.master.master_data[cost_bl_index[i]].data[
                            project_name
                        ][key[0]]
                        if rdel is None:
                            rdel = 0

                        cdel = self.master.master_data[cost_bl_index[i]].data[
                            project_name
                        ][key[1]]
                        if cdel is None:
                            cdel = 0

                        ngov = self.master.master_data[cost_bl_index[i]].data[
                            project_name
                        ][key[2]]
                        if ngov is None:
                            ngov = 0

                        total = round(rdel + cdel + ngov)
                        group_total += total
                    except TypeError:  # handle None types, which are present if project not reporting last quarter.
                        rdel = 0
                        cdel = 0
                        ngov = 0
                        total = 0
                        group_total += total

                    if i == 0:  # current quarter
                        if x == 0:  # spent
                            try:  # handling for spend to date figures which are not present in all masters
                                rdel_std = self.master.master_data[
                                    cost_bl_index[i]
                                ].data[project_name]["20-21 RDEL STD one off new costs"]
                                if rdel_std is None:
                                    rdel_std = 0
                                cdel_std = self.master.master_data[
                                    cost_bl_index[i]
                                ].data[project_name]["20-21 CDEL STD one off new costs"]
                                if cdel_std is None:
                                    cdel_std = 0
                                ngov_std = self.master.master_data[
                                    cost_bl_index[i]
                                ].data[project_name]["20-21 CDEL STD Non Gov costs"]
                                if ngov_std is None:
                                    ngov_std = 0
                                group_rdel_spent += round(rdel + rdel_std)
                                group_cdel_spent += round(cdel + cdel_std)
                                group_ngov_spent += round(ngov + ngov_std)
                            except KeyError:
                                group_rdel_spent += rdel
                                group_cdel_spent += cdel
                                group_ngov_spent += ngov
                        if x == 1:  # profiled
                            group_rdel_profiled += rdel
                            group_cdel_profiled += cdel
                            group_ngov_profiled += ngov
                        if x == 2:  # unprofiled
                            group_rdel_unprofiled += rdel
                            group_cdel_unprofiled += cdel
                            group_ngov_unprofiled += ngov

                if x == 0:  # spent
                    try:  # handling for spend to date figures which are not present in all masters
                        rdel_std = self.master.master_data[cost_bl_index[i]].data[
                            project_name
                        ]["20-21 RDEL STD one off new costs"]
                        cdel_std = self.master.master_data[cost_bl_index[i]].data[
                            project_name
                        ]["20-21 CDEL STD one off new costs"]
                        ngov_std = self.master.master_data[cost_bl_index[i]].data[
                            project_name
                        ]["20-21 CDEL STD Non Gov costs"]
                        std_list = [
                            rdel_std,
                            cdel_std,
                            ngov_std,
                        ]  # converts none types to zero
                        for s, std in enumerate(std_list):
                            if std is None:
                                std_list[s] = 0
                        spent.append(round(group_total + sum(std_list)))
                    except (
                        KeyError,
                        TypeError,
                    ):  # Note. TypeError here as projects may have no baseline
                        spent.append(group_total)
                if x == 1:  # profiled
                    profiled.append(group_total)
                if x == 2:  # unprofiled
                    unprofiled.append(group_total)

        cat_spent = [group_rdel_spent, group_cdel_spent, group_ngov_spent]
        cat_profiled = [group_rdel_profiled, group_cdel_profiled, group_ngov_profiled]
        cat_unprofiled = [
            group_rdel_unprofiled,
            group_cdel_unprofiled,
            group_ngov_unprofiled,
        ]
        final_cat_profiled = calculate_profiled(cat_profiled, cat_spent, cat_unprofiled)

        all_profiled = calculate_profiled(profiled, spent, unprofiled)

        self.cat_spent = cat_spent
        self.cat_profiled = final_cat_profiled
        self.cat_unprofiled = cat_unprofiled
        self.spent = spent
        self.profiled = all_profiled
        self.unprofiled = unprofiled
        self.y_scale_max = max(profiled)
        self.entity = group

    def get_cost_profile(self, group: List[str] or str, baseline: str) -> None:
        """Returns several lists which contain the sum of different cost profiles for the group of project
        contained with the master"""
        self.group = group
        self.baseline = baseline

        current_profile = []
        last_profile = []
        baseline_profile_one = []
        baseline_profile_two = []
        baseline_profile_three = []
        rdel_current_profile = []
        cdel_current_profile = []
        ngov_current_profile = []
        missing_projects = []

        group = string_conversion(self.group)

        for i in range(5):
            yearly_profile = []
            rdel_yearly_profile = []
            cdel_yearly_profile = []
            ngov_yearly_profile = []
            for year in YEAR_LIST:
                cost_total = 0
                rdel_total = 0
                cdel_total = 0
                ngov_total = 0
                for cost_type in COST_KEY_LIST:
                    for project_name in group:
                        project_bl_index = self.master.bl_index[baseline][project_name]
                        try:
                            cost = self.master.master_data[project_bl_index[i]].data[
                                project_name
                            ][year + cost_type]
                            if cost is None:
                                cost = 0
                            cost_total += cost
                        except KeyError:  # to handle data across different financial years
                            cost = 0
                            cost_total += cost
                        except TypeError:  # Handles projects not present in the previous quarter
                            missing_projects.append(
                                str(project_name)
                            )  # projects added here. message is below.
                            cost = 0
                            cost_total += cost
                        except IndexError:  # Handles project baseline index
                            # TODO improve this loop
                            if i == 3:
                                try:
                                    cost = self.master.master_data[
                                        project_bl_index[2]
                                    ].data[project_name][year + cost_type]
                                    if cost is None:
                                        cost = 0
                                    cost_total += cost
                                except KeyError:  # to handle data across different financial years
                                    cost = 0
                                    cost_total += cost
                            if i == 4:
                                try:
                                    cost = self.master.master_data[
                                        project_bl_index[3]
                                    ].data[project_name][year + cost_type]
                                    if cost is None:
                                        cost = 0
                                    cost_total += cost
                                except KeyError:  # to handle data across different financial years
                                    cost = 0
                                    cost_total += cost
                                except IndexError:
                                    try:
                                        cost = self.master.master_data[
                                            project_bl_index[2]
                                        ].data[project_name][year + cost_type]
                                        if cost is None:
                                            cost = 0
                                        cost_total += cost
                                    except KeyError:  # to handle data across different financial years
                                        cost = 0
                                        cost_total += cost

                        if cost_type == COST_KEY_LIST[0]:  # rdel
                            rdel_total += cost
                        if cost_type == COST_KEY_LIST[1]:  # cdel
                            cdel_total += cost
                        if cost_type == COST_KEY_LIST[2]:  # ngov
                            ngov_total += cost

                yearly_profile.append(round(cost_total))
                rdel_yearly_profile.append(round(rdel_total))
                cdel_yearly_profile.append(round(cdel_total))
                ngov_yearly_profile.append(round(ngov_total))

            if i == 0:
                current_profile = yearly_profile
                rdel_current_profile = rdel_yearly_profile
                cdel_current_profile = cdel_yearly_profile
                ngov_current_profile = ngov_yearly_profile
            if i == 1:
                last_profile = yearly_profile
            if i == 2:
                baseline_profile_one = yearly_profile
            if i == 3:
                baseline_profile_two = yearly_profile
            if i == 4:
                baseline_profile_three = yearly_profile

        missing_projects = list(set(missing_projects))  # if TypeError raised above
        if len(missing_projects) is not 0:
            print(
                "NOTE: The following project(s) were not part of the portfolio last quarter "
                + str(missing_projects)
                + " this means current quarter and last quarter cost profiles are not like for like."
                " If you would like a like for like comparison between current and last quarter"
                " remove this project(s) from the master group."
            )

        self.current_profile = current_profile
        self.last_profile = last_profile
        self.baseline_profile_one = baseline_profile_one
        self.baseline_profile_two = baseline_profile_two
        self.baseline_profile_three = baseline_profile_three
        self.rdel_profile = rdel_current_profile
        self.cdel_profile = cdel_current_profile
        self.ngov_profile = ngov_current_profile
        self.entity = group


class BenefitsData:
    def __init__(self, master: Master):
        self.master = master
        self.cat_delivered = []
        self.cat_profiled = []
        self.cat_unprofiled = []
        self.delivered = []
        self.profiled = []
        self.unprofiled = []
        self.y_scale_max = 0
        self.y_scale_min = 0
        self.economic_max = 0

    def get_ben_totals(self, group: List[str] or str, baseline: str) -> None:
        """Returns lists containing the sum total of group (of projects) benefits,
        sliced in different ways. Cumbersome for loop used at the moment, but
        is the least cumbersome loop I could design!"""
        self.group = group
        self.baseline = baseline

        delivered = []
        profiled = []
        unprofiled = []
        group_cash_dev = 0
        group_uncash_dev = 0
        group_economic_dev = 0
        group_disben_dev = 0
        group_cash_profiled = 0
        group_uncash_profiled = 0
        group_economic_profiled = 0
        group_disben_profiled = 0
        group_cash_unprofiled = 0
        group_uncash_unprofiled = 0
        group_economic_unprofiled = 0
        group_disben_unprofiled = 0

        group = string_conversion(self.group)

        # where to store this function. need it at a global level for CostData Class
        def calculate_profiled(p: List[int], s: List[int], unpro: List[int]) -> list:
            """small helper function to calculate the proper profiled amount. This is necessary as
            other wise 'profiled' would actually be the total figure.
            p = profiled list
            s = spent list
            unpro = unprofiled list"""
            f_profiled = []
            for y, amount in enumerate(p):
                t = amount - (s[y] + unpro[y])
                f_profiled.append(t)
            return f_profiled

        for i in range(3):
            for x, key in enumerate(BEN_TYPE_KEY_LIST):
                group_total = 0
                for project in group:
                    ben_bl_index = self.master.bl_index[baseline][project]
                    try:
                        cash = round(
                            self.master.master_data[ben_bl_index[i]].data[project][
                                key[0]
                            ]
                        )
                        uncash = round(
                            self.master.master_data[ben_bl_index[i]].data[project][
                                key[1]
                            ]
                        )
                        economic = round(
                            self.master.master_data[ben_bl_index[i]].data[project][
                                key[2]
                            ]
                        )
                        disben = round(
                            self.master.master_data[ben_bl_index[i]].data[project][
                                key[3]
                            ]
                        )

                        total = round(cash + uncash + economic + disben)
                        group_total += total
                    except TypeError:  # handle None types, which are present if project not reporting last quarter.
                        cash = 0
                        uncash = 0
                        economic = 0
                        disben = 0
                        total = 0
                        group_total += total

                    if i == 0:  # current quarter
                        if x == 0:  # spent
                            group_cash_dev += cash
                            group_uncash_dev += uncash
                            group_economic_dev += economic
                            group_disben_dev += disben
                        if x == 1:  # profiled
                            group_cash_profiled += cash
                            group_uncash_profiled += uncash
                            group_economic_profiled += economic
                            group_disben_profiled += disben
                        if x == 2:  # unprofiled
                            group_cash_unprofiled += cash
                            group_uncash_unprofiled += uncash
                            group_economic_unprofiled += economic
                            group_disben_unprofiled += disben

                if x == 0:  # spent
                    delivered.append(group_total)
                if x == 1:  # profiled
                    profiled.append(group_total)
                if x == 2:  # unprofiled
                    unprofiled.append(group_total)

        cat_spent = [
            group_cash_dev,
            group_uncash_dev,
            group_economic_dev,
            group_disben_dev,
        ]
        cat_profiled = [
            group_cash_profiled,
            group_uncash_profiled,
            group_economic_profiled,
            group_disben_profiled,
        ]
        cat_unprofiled = [
            group_cash_unprofiled,
            group_uncash_unprofiled,
            group_economic_unprofiled,
            group_disben_unprofiled,
        ]
        final_cat_profiled = calculate_profiled(cat_profiled, cat_spent, cat_unprofiled)
        all_profiled = calculate_profiled(profiled, delivered, unprofiled)

        self.cat_delivered = cat_spent
        self.cat_profiled = final_cat_profiled
        self.cat_unprofiled = cat_unprofiled
        self.delivered = delivered
        self.profiled = all_profiled
        self.unprofiled = unprofiled
        self.y_scale_max = max(profiled)
        self.y_scale_min = min(
            [group_disben_dev, group_disben_profiled, group_disben_unprofiled]
        )
        self.economic_max = max(
            [group_economic_dev, group_economic_unprofiled, group_economic_profiled]
        )


def vfm_matplotlib_graph(labels, current_qrt, last_qrt, title):
    #  Need to split this strings over two lines on x axis
    for n, i in enumerate(labels):
        if i == "Very High and Financially Positive":
            labels[n] = "Very High and \n Financially Positive"
        if i == "Economically Positive":
            labels[n] = "Economically \n Positive"

    x = np.arange(len(labels))  # the label locations
    width = 0.35  # the width of the bars

    fig, ax = plt.subplots()
    rects_one = ax.bar(x - width / 2, current_qrt, width, label="This quarter")
    rects_two = ax.bar(x + width / 2, last_qrt, width, label="Last quarter")

    # Add some text for labels, title and custom x-axis tick labels, etc.
    # ax.set_ylabel('Number')
    ax.set_title(title)
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    # Rotate the tick labels and set their alignment.
    # plt.setp(ax.get_xticklabels(), alignment=)
    ax.legend()

    def autolabel(rects):
        """Attach a text label above each bar in *rects*, displaying its height."""
        for rect in rects:
            height = rect.get_height()
            ax.annotate(
                "{}".format(height),
                xy=(rect.get_x() + rect.get_width() / 2, height),
                xytext=(0, 3),  # 3 points vertical offset
                textcoords="offset points",
                ha="center",
                va="bottom",
            )

    autolabel(rects_one)
    autolabel(rects_two)

    fig.tight_layout()

    fig.savefig(root_path / "output/{}.png".format(title), bbox_inches="tight")

    plt.show()


def set_figure_size(graph_type: str) -> Tuple[int, int]:
    if graph_type == "half horizontal":
        return 11.69, 4.10
    if graph_type == "full horizontal":
        return 11.69, 8.20


def cost_profile_graph(
    fig_size: str, cost_master: CostData, *args: Tuple[Optional[str]]
):
    """Compiles a matplotlib line chart for costs of GROUP of projects contained within cost_master class.
    As as default last quarters profile is not included. It creates two plots. First plot shows overall
    profile in current, last quarters anb baseline form. Second plot shows rdel, cdel, and 'non-gov' cost profile"""

    fig, (ax1, ax2) = plt.subplots(2)  # two subplots for this chart
    fig.set_size_inches(set_figure_size(fig_size))

    """cost profile charts"""
    if len(cost_master.entity) == 1:
        fig.suptitle(cost_master.entity[0] + " Cost Profile", fontweight="bold")
    else:
        try:
            fig.suptitle(args[0] + " Cost Profile", fontweight="bold")  # title
        except IndexError:
            print("You need to provide a title for this chart")

    # Overall cost profile chart
    if (
        sum(cost_master.baseline_profile_one) != 0
    ):  # handling in the event that group of projects have no baseline profile.
        ax1.plot(
            YEAR_LIST,
            np.array(cost_master.baseline_profile_one),  # baseline profile
            label="Baseline",
            linewidth=3.0,
            marker="o",
        )
    else:
        pass
    if (
        sum(cost_master.last_profile) != 0
    ):  # handling in the event that group of projects have no last quarter profile
        ax1.plot(
            YEAR_LIST,
            np.array(cost_master.last_profile),  # last quarter profile
            label="Last quarter",
            linewidth=3.0,
            marker="o",
        )
    else:
        pass
    ax1.plot(
        YEAR_LIST,
        np.array(cost_master.current_profile),  # current profile
        label="Latest",
        linewidth=3.0,
        marker="o",
    )

    # Chart styling
    ax1.tick_params(axis="x", which="major", labelsize=6, rotation=45)
    ax1.set_ylabel("Cost (£m)")
    ylab1 = ax1.yaxis.get_label()
    ylab1.set_style("italic")
    ylab1.set_size(8)
    ax1.grid(color="grey", linestyle="-", linewidth=0.2)
    ax1.legend(prop={"size": 6})
    ax1.set_title(
        "Fig 1 - cost profile changes", loc="left", fontsize=8, fontweight="bold"
    )

    # plot rdel, cdel, non-gov chart data
    if (
        sum(cost_master.ngov_profile) != 0
    ):  # if statement as most projects don't have ngov cost.
        ax2.plot(
            YEAR_LIST,
            np.array(cost_master.ngov_profile),
            label="Non-Gov",
            linewidth=3.0,
            marker="o",
        )
    ax2.plot(
        YEAR_LIST,
        np.array(cost_master.cdel_profile),
        label="CDEL",
        linewidth=3.0,
        marker="o",
    )
    ax2.plot(
        YEAR_LIST,
        np.array(cost_master.rdel_profile),
        label="RDEL",
        linewidth=3.0,
        marker="o",
    )

    # rdel/cdel profile chart styling
    ax2.tick_params(axis="x", which="major", labelsize=6, rotation=45)
    ax2.set_xlabel("Financial Years")
    ax2.set_ylabel("Cost (£m)")
    xlab2 = ax2.xaxis.get_label()
    ylab2 = ax2.yaxis.get_label()
    xlab2.set_style("italic")
    xlab2.set_size(8)
    ylab2.set_style("italic")
    ylab2.set_size(8)
    ax2.grid(color="grey", linestyle="-", linewidth=0.2)
    ax2.legend(prop={"size": 6})
    ax2.set_title(
        "Fig 2 - current cost type profile", loc="left", fontsize=8, fontweight="bold"
    )

    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # size/fit of chart

    plt.show()

    return fig


# what does this output plt.figure.Figure?
def cost_profile_baseline_graph(cost_master: CostData, *args: Tuple[Optional[str]]):
    """Compiles a matplotlib line chart for costs of GROUP of projects contained within cost_master class.
    As as default last quarters profile is not included. It creates two plots. First plot shows overall
    profile in current, last quarters anb baseline form. Second plot shows rdel, cdel, and 'non-gov' cost profile"""

    fig, (ax1, ax2) = plt.subplots(2)  # two subplots for this chart

    """cost profile charts"""
    if len(cost_master.entity) == 1:
        fig.suptitle(cost_master.entity[0] + " Cost Profile", fontweight="bold")
    else:
        fig.suptitle(args[0] + " Cost Profile", fontweight="bold")  # title

    # Overall cost profile chart
    if (
        sum(cost_master.baseline_profile_three) != 0
    ):  # handling in the event that group of projects have no baseline profile.
        ax1.plot(
            YEAR_LIST,
            np.array(cost_master.baseline_profile_three),  # baseline profile
            label="Baseline 3",
            linewidth=3.0,
            marker="o",
        )
    else:
        pass
    if (
        sum(cost_master.baseline_profile_two) != 0
    ):  # handling in the event that group of projects have no baseline profile.
        ax1.plot(
            YEAR_LIST,
            np.array(cost_master.baseline_profile_two),  # baseline profile
            label="Baseline 2",
            linewidth=3.0,
            marker="o",
        )
    else:
        pass
    if (
        sum(cost_master.baseline_profile_one) != 0
    ):  # handling in the event that group of projects have no last quarter profile
        ax1.plot(
            YEAR_LIST,
            np.array(cost_master.baseline_profile_one),  # last quarter profile
            label="Baseline 1",
            linewidth=3.0,
            marker="o",
        )
    else:
        pass
    ax1.plot(
        YEAR_LIST,
        np.array(cost_master.current_profile),  # current profile
        label="Latest",
        linewidth=3.0,
        marker="o",
    )

    # Chart styling
    ax1.tick_params(axis="x", which="major", labelsize=6, rotation=45)
    ax1.set_ylabel("Cost (£m)")
    ylab1 = ax1.yaxis.get_label()
    ylab1.set_style("italic")
    ylab1.set_size(8)
    ax1.grid(color="grey", linestyle="-", linewidth=0.2)
    ax1.legend(prop={"size": 6})
    ax1.set_title(
        "Fig 1 - cost profile changes", loc="left", fontsize=8, fontweight="bold"
    )

    # plot rdel, cdel, non-gov chart data
    if (
        sum(cost_master.ngov_profile) != 0
    ):  # if statement as most projects don't have ngov cost.
        ax2.plot(
            YEAR_LIST,
            np.array(cost_master.ngov_profile),
            label="Non-Gov",
            linewidth=3.0,
            marker="o",
        )
    ax2.plot(
        YEAR_LIST,
        np.array(cost_master.cdel_profile),
        label="CDEL",
        linewidth=3.0,
        marker="o",
    )
    ax2.plot(
        YEAR_LIST,
        np.array(cost_master.rdel_profile),
        label="RDEL",
        linewidth=3.0,
        marker="o",
    )

    # rdel/cdel profile chart styling
    ax2.tick_params(axis="x", which="major", labelsize=6, rotation=45)
    ax2.set_xlabel("Financial Years")
    ax2.set_ylabel("Cost (£m)")
    xlab2 = ax2.xaxis.get_label()
    ylab2 = ax2.yaxis.get_label()
    xlab2.set_style("italic")
    xlab2.set_size(8)
    ylab2.set_style("italic")
    ylab2.set_size(8)
    ax2.grid(color="grey", linestyle="-", linewidth=0.2)
    ax2.legend(prop={"size": 6})
    ax2.set_title(
        "Fig 2 - current cost type profile", loc="left", fontsize=8, fontweight="bold"
    )

    plt.show()

    return fig


def spent_calculation(
    master: Dict[str, Union[str, date, int, float]], project: str
) -> int:
    keys = [
        "Pre-profile RDEL",
        "20-21 RDEL STD Total",
        "Pre-profile CDEL Forecast one off new costs",
        "20-21 CDEL STD Total",
        "Pre-profile Forecast Non-Gov",
        "20-21 CDEL STD Non Gov costs",
    ]

    total = 0
    for k in keys:
        try:
            total += master.data[project][k]
        except TypeError:  # None types
            pass

    return total


def open_word_doc(wd_path: str) -> Document:
    """Function stores an empty word doc as a variable"""
    return Document(wd_path)


def wd_heading(
    doc: Document, project_info: Dict[str, Union[str, int]], project_name: str
) -> None:
    """Function adds header to word doc"""
    font = doc.styles["Normal"].font
    font.name = "Arial"
    font.size = Pt(12)

    heading = str(
        project_info.data[project_name]["Abbreviations"]
    )  # integrate into master
    intro = doc.add_heading(str(heading), 0)
    intro.alignment = 1
    intro.bold = True


def key_contacts(doc: Document, master: Master, project_name: str) -> None:
    """Function adds key contact details"""
    sro_name = master.master_data[0].data[project_name][
        "Senior Responsible Owner (SRO)"
    ]
    if sro_name is None:
        sro_name = "tbc"

    sro_email = master.master_data[0].data[project_name][
        "Senior Responsible Owner (SRO) - Email"
    ]
    if sro_email is None:
        sro_email = "email: tbc"

    sro_phone = master.master_data[0].data[project_name]["SRO Phone No."]
    if sro_phone == None:
        sro_phone = "phone number: tbc"

    doc.add_paragraph(
        "SRO: " + str(sro_name) + ", " + str(sro_email) + ", " + str(sro_phone)
    )

    pd_name = master.master_data[0].data[project_name]["Project Director (PD)"]
    if pd_name is None:
        pd_name = "TBC"

    pd_email = master.master_data[0].data[project_name]["Project Director (PD) - Email"]
    if pd_email is None:
        pd_email = "email: tbc"

    pd_phone = master.master_data[0].data[project_name]["PD Phone No."]
    if pd_phone is None:
        pd_phone = "TBC"

    doc.add_paragraph(
        "PD: " + str(pd_name) + ", " + str(pd_email) + ", " + str(pd_phone)
    )

    contact_name = master.master_data[0].data[project_name]["Working Contact Name"]
    if contact_name is None:
        contact_name = "TBC"

    contact_email = master.master_data[0].data[project_name]["Working Contact Email"]
    if contact_email is None:
        contact_email = "email: tbc"

    contact_phone = master.master_data[0].data[project_name][
        "Working Contact Telephone"
    ]
    if contact_phone is None:
        contact_phone = "TBC"

    doc.add_paragraph(
        "PfM reporting lead: "
        + str(contact_name)
        + ", "
        + str(contact_email)
        + ", "
        + str(contact_phone)
    )


def dca_table(doc: Document, master: Master, project_name: str) -> None:
    """Creates SRO confidence table"""
    w_table = doc.add_table(rows=1, cols=5)
    hdr_cells = w_table.rows[0].cells
    hdr_cells[0].text = "Delivery confidence"
    hdr_cells[1].text = "This quarter"
    hdr_cells[2].text = str(master.master_data[1].quarter)
    hdr_cells[3].text = str(master.master_data[2].quarter)
    hdr_cells[4].text = str(master.master_data[3].quarter)

    for x, dca_key in enumerate(SRO_CONF_KEY_LIST):
        row_cells = w_table.add_row().cells
        row_cells[0].text = dca_key
        for i, m in enumerate(master.master_data[:4]):  # last four masters taken
            try:
                rating = convert_rag_text(m.data[project_name][dca_key])
                row_cells[i + 1].text = rating
                cell_colouring(row_cells[i + 1], rating)
            except (KeyError, TypeError):
                row_cells[i + 1].text = "N/A"

    w_table.style = "Table Grid"
    make_rows_bold([w_table.rows[0]])  # makes top of table bold.
    # make_columns_bold([table.columns[0]]) #right cells in table bold
    column_widths = (Cm(3.9), Cm(2.9), Cm(2.9), Cm(2.9), Cm(2.9))
    set_col_widths(w_table, column_widths)


def dca_narratives(doc: Document, master: Master, project_name: str) -> None:
    """Places all narratives into document and checks for differences between
    current and last quarter"""

    doc.add_paragraph()
    p = doc.add_paragraph()
    text = "*Red text highlights changes in narratives from last quarter"
    p.add_run(text).font.color.rgb = RGBColor(255, 0, 0)

    headings_list = [
        "SRO delivery confidence narrative",
        "Financial cost narrative",
        "Financial comparison with last quarter",
        "Financial comparison with baseline",
        "Benefits Narrative",
        "Benefits comparison with last quarter",
        "Benefits comparison with baseline",
        "Milestone narrative",
    ]

    narrative_keys_list = [
        "Departmental DCA Narrative",
        "Project Costs Narrative",
        "Cost comparison with last quarters cost narrative",
        "Cost comparison within this quarters cost narrative",
        "Benefits Narrative",
        "Ben comparison with last quarters cost - narrative",
        "Ben comparison within this quarters cost - narrative",
        "Milestone Commentary",
    ]

    for x in range(len(headings_list)):
        doc.add_paragraph().add_run(str(headings_list[x])).bold = True
        text_one = str(master.master_data[0].data[project_name][narrative_keys_list[x]])
        try:
            text_two = str(
                master.master_data[1].data[project_name][narrative_keys_list[x]]
            )
        except KeyError:
            text_two = text_one

        # There are two options here for comparing text. Have left this for now.
        # compare_text_showall(dca_a, dca_b, doc)
        compare_text_new_and_old(text_one, text_two, doc)


def change_word_doc_landscape(doc: Document) -> Document():
    new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)  # new page
    # change to landscape
    new_width, new_height = new_section.page_height, new_section.page_width
    new_section.orientation = WD_ORIENTATION.LANDSCAPE
    new_section.page_width = new_width
    new_section.page_height = new_height
    return doc


def put_matplotlib_fig_into_word(doc: Document, fig) -> None:
    """Places line graph cost profile into word document"""
    # Place fig in word doc.
    fig.savefig("cost_profile.png")
    doc.add_picture("cost_profile.png", width=Inches(8))  # to place nicely in doc
    os.remove("cost_profile.png")
    plt.close()  # automatically closes figure so don't need to do manually.


def convert_rag_text(dca_rating: str) -> None:
    """Converts RAG name into a acronym"""

    if dca_rating == "Green":
        return "G"
    elif dca_rating == "Amber/Green":
        return "A/G"
    elif dca_rating == "Amber":
        return "A"
    elif dca_rating == "Amber/Red":
        return "A/R"
    elif dca_rating == "Red":
        return "R"
    else:
        return ""


def cell_colouring(word_table_cell: table.Table.cell, colour: str) -> None:
    """Function that handles cell colouring for word documents"""

    try:
        if colour == "R":
            colour = parse_xml(r'<w:shd {} w:fill="cb1f00"/>'.format(nsdecls("w")))
        elif colour == "A/R":
            colour = parse_xml(r'<w:shd {} w:fill="f97b31"/>'.format(nsdecls("w")))
        elif colour == "A":
            colour = parse_xml(r'<w:shd {} w:fill="fce553"/>'.format(nsdecls("w")))
        elif colour == "A/G":
            colour = parse_xml(r'<w:shd {} w:fill="a5b700"/>'.format(nsdecls("w")))
        elif colour == "G":
            colour = parse_xml(r'<w:shd {} w:fill="17960c"/>'.format(nsdecls("w")))

        word_table_cell._tc.get_or_add_tcPr().append(colour)

    except TypeError:
        pass


def make_rows_bold(rows: list) -> None:
    """This function makes text bold in a list of row numbers for a word document"""
    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True


def set_col_widths(word_table: table, widths: list) -> None:
    """This function sets the width of table in a word document"""
    for row in word_table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width


def compare_text_new_and_old(text_1: str, text_2: str, doc: Document) -> None:
    """compares two sets of text and highlights differences in red text."""

    comp = difflib.Differ()
    diff = list(comp.compare(text_2.split(), text_1.split()))
    new_text = diff
    y = doc.add_paragraph()

    for i in range(0, len(diff)):
        f = len(diff) - 1
        if i < f:
            a = i - 1
        else:
            a = i

        if diff[i][0:3] == "  |":
            j = i + 1
            if diff[i][0:3] and diff[a][0:3] == "  |":
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == "+ |":
            if diff[i][0:3] and diff[a][0:3] == "+ |":
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == "- |":
            pass
        elif diff[i][0:3] == "  -":
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0:3] == "  •":
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0] == "+":
            w = len(diff[i])
            g = diff[i][1:w]
            y.add_run(g).font.color.rgb = RGBColor(255, 0, 0)
        elif diff[i][0] == "-":
            pass
        elif diff[i][0] == "?":
            pass
        else:
            if diff[i] != "+ |":
                y.add_run(diff[i][1:])


def make_file_friendly(quarter_str: str) -> str:
    """Converts datamaps.api project_data_from_master quarter data into a string to use when
    saving output files. Courtesy of M Lemon."""
    regex = r"Q(\d) (\d+)\/(\d+)"
    return re.sub(regex, r"Q\1_\2_\3", quarter_str)


def total_costs_benefits_bar_chart(
    fig_size: str,
    cost_master: CostData,
    ben_master: BenefitsData,
    *args: Tuple[Optional[str]]
) -> plt.figure:
    """compiles a matplotlib bar chart which shows total project costs"""
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2)  # four sub plots
    fig.set_size_inches(set_figure_size(fig_size))

    if len(cost_master.entity) == 1:
        fig.suptitle(
            cost_master.entity[0] + " Cost and Benefits Profile", fontweight="bold"
        )
    else:
        fig.suptitle(args[0] + " Cost and Benefits Profile", fontweight="bold")  # title

    # Y AXIS SCALE MAX
    highest_int = max([cost_master.y_scale_max, ben_master.y_scale_max])
    y_max = highest_int + percentage(5, highest_int)
    ax1.set_ylim(0, y_max)

    # COST SPENT, PROFILED AND UNPROFILED
    labels = ["Latest", "Last Quarter", "Baseline"]
    width = 0.5
    ax1.bar(labels, np.array(cost_master.spent), width, label="Spent")
    ax1.bar(
        labels,
        np.array(cost_master.profiled),
        width,
        bottom=np.array(cost_master.spent),
        label="Profiled",
    )
    ax1.bar(
        labels,
        np.array(cost_master.unprofiled),
        width,
        bottom=np.array(cost_master.spent) + np.array(cost_master.profiled),
        label="Unprofiled",
    )
    ax1.legend(prop={"size": 6})
    ax1.set_ylabel("Cost (£m)")
    ylab1 = ax1.yaxis.get_label()
    ylab1.set_style("italic")
    ylab1.set_size(8)
    ax1.tick_params(axis="x", which="major", labelsize=6)
    ax1.tick_params(axis="y", which="major", labelsize=6)
    ax1.set_title(
        "Fig 1 - Total costs: change over time",
        loc="left",
        fontsize=8,
        fontweight="bold",
    )

    # RDEL, CDEL AND NON-GOV TOTALS BAR CHART
    labels = ["RDEL", "CDEL", "Non Gov"]
    width = 0.5
    ax2.bar(labels, np.array(cost_master.cat_spent), width, label="Spent")
    ax2.bar(
        labels,
        np.array(cost_master.cat_profiled),
        width,
        bottom=np.array(cost_master.cat_spent),
        label="Profiled",
    )
    ax2.bar(
        labels,
        np.array(cost_master.cat_unprofiled),
        width,
        bottom=np.array(cost_master.cat_spent) + np.array(cost_master.cat_profiled),
        label="Unprofiled",
    )
    ax2.legend(prop={"size": 6})
    ax2.set_ylabel("Costs (£m)")
    ylab3 = ax2.yaxis.get_label()
    ylab3.set_style("italic")
    ylab3.set_size(8)
    ax2.tick_params(axis="x", which="major", labelsize=6)
    ax2.tick_params(axis="y", which="major", labelsize=6)
    ax2.set_title(
        "Fig 2 - wlc cost type break down", loc="left", fontsize=8, fontweight="bold"
    )

    ax2.set_ylim(0, y_max)  # scale y axis max

    # BENEFITS SPENT, PROFILED AND UNPROFILED
    labels = ["Latest", "Last Quarter", "Baseline"]
    width = 0.5
    ax3.bar(labels, np.array(ben_master.delivered), width, label="Delivered")
    ax3.bar(
        labels,
        np.array(ben_master.profiled),
        width,
        bottom=np.array(ben_master.delivered),
        label="Profiled",
    )
    ax3.bar(
        labels,
        np.array(ben_master.unprofiled),
        width,
        bottom=np.array(ben_master.delivered) + np.array(ben_master.profiled),
        label="Unprofiled",
    )
    ax3.legend(prop={"size": 6})
    ax3.set_ylabel("Benefits (£m)")
    ylab3 = ax3.yaxis.get_label()
    ylab3.set_style("italic")
    ylab3.set_size(8)
    ax3.tick_params(axis="x", which="major", labelsize=6)
    ax3.tick_params(axis="y", which="major", labelsize=6)
    ax3.set_title(
        "Fig 3 - ben total change over time", loc="left", fontsize=8, fontweight="bold"
    )

    ax3.set_ylim(0, y_max)

    # BENEFITS CASHABLE, NON-CASHABLE, ECONOMIC, DISBENEFIT BAR CHART
    labels = ["Cashable", "Non-Cashable", "Economic", "Disbenefit"]
    width = 0.5
    ax4.bar(labels, np.array(ben_master.cat_delivered), width, label="Delivered")
    ax4.bar(
        labels,
        np.array(ben_master.cat_profiled),
        width,
        bottom=np.array(ben_master.cat_delivered),
        label="Profiled",
    )
    ax4.bar(
        labels,
        np.array(ben_master.cat_unprofiled),
        width,
        bottom=np.array(ben_master.cat_delivered) + np.array(ben_master.cat_profiled),
        label="Unprofiled",
    )
    ax4.legend(prop={"size": 6})
    ax4.set_ylabel("Benefits (£m)")
    ylab4 = ax4.yaxis.get_label()
    ylab4.set_style("italic")
    ylab4.set_size(8)
    ax4.tick_params(axis="x", which="major", labelsize=6)
    ax4.tick_params(axis="y", which="major", labelsize=6)
    ax4.set_title(
        "Fig 4 - benefits profile type", loc="left", fontsize=8, fontweight="bold"
    )

    y_min = ben_master.y_scale_min + percentage(40, ben_master.y_scale_min)
    if ben_master.economic_max > y_max:
        ax4.set_ylim(y_min, y_max)
    else:
        ax4.set_ylim(
            y_min, y_max
        )  # possible for economic benefits figure to be the largest

    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # size/fit of chart

    plt.show()  # doesn't seem to do anything

    return fig


def check_baselines(master: Master) -> None:
    """checks that projects have the correct baseline information. stops the
    programme if baselines are missing"""

    for v in BASELINE_TYPES.values():
        for p in master.current_projects:
            baselines = master.bl_index[v][p]
            if len(baselines) <= 2:
                print(
                    p
                    + " does not have a baseline point for "
                    + v
                    + " this could cause the programme to"
                    "crash. Therefore the programme is stopping. "
                    "Please amend the data for " + p + " so that "
                    " it has at least one baseline point for " + v
                )
                break
        else:
            continue
        break


def percentage(percent: int, whole: float) -> int:
    return round((percent * whole) / 100.0)


def get_old_fy_cost_data(
    master_file: typing.TextIO, project_id_wb: typing.TextIO
) -> None:
    """
    Gets all old financial data from a specified master and places into project id document.
    """
    master = project_data_from_master(
        master_file, 1, 2010
    )  # random year specified as not in use
    wb = load_workbook(project_id_wb)
    ws = wb.active

    for i in range(1, ws.max_column + 1):
        project_name = ws.cell(row=1, column=1 + i).value
        for row_num in range(2, ws.max_row + 1):
            key = ws.cell(row=row_num, column=1).value
            try:
                if key in master.data[project_name].keys():
                    ws.cell(row=row_num, column=1 + i).value = master.data[
                        project_name
                    ][key]
            except KeyError:  # project might not be present in quarter
                pass

    wb.save(project_id_wb)


def run_get_old_fy_data(master_files_list: list, project_id_wb: typing.TextIO) -> None:
    for f in reversed(
        master_files_list
    ):  # reversed so it gets the latest data in masters
        get_old_fy_cost_data(f, project_id_wb)


def place_old_fy_data_into_master_wb(
    master_file: typing.TextIO, project_id_wb: typing.TextIO
) -> None:
    """
    places all old financial year data into master files.
    """
    id_master = project_data_from_master(
        project_id_wb, 2, 2020
    )  # random year specify as not used
    wb = load_workbook(master_file)
    ws = wb.active

    for i in range(1, ws.max_column + 1):
        project_name = ws.cell(row=1, column=1 + i).value
        for row_num in range(2, ws.max_row + 1):
            key = ws.cell(row=row_num, column=1).value
            try:
                if key in id_master.data[project_name].keys():
                    ws.cell(row=row_num, column=1 + i).value = id_master.data[
                        project_name
                    ][key]
            except KeyError:  # project might not be present in quarter
                pass

    wb.save(master_file)


def run_place_old_fy_data_into_masters(
    master_files_list: list, project_id_wb: typing.TextIO
) -> None:
    for f in master_files_list:
        place_old_fy_data_into_master_wb(f, project_id_wb)


def put_key_change_master_into_dict(key_change_file: typing.TextIO) -> Dict[str, str]:
    """
    places key information i.e. keys old and new names from wb into a python dictionary
    """
    wb = load_workbook(key_change_file)
    ws = wb.active

    output_dict = {}
    for x in range(1, ws.max_row + 1):
        key = ws.cell(row=x, column=1).value
        codename = ws.cell(row=x, column=2).value
        output_dict[key] = codename

    return output_dict


def alter_wb_master_file_key_names(
    master_file: typing.TextIO, key_change_dict: Dict[str, str]
) -> workbook:
    """
    places altered key names, from the key change master dictionary, into master wb(s).
    """
    wb = load_workbook(master_file)
    ws = wb.active

    for row_num in range(2, ws.max_row + 1):
        for (
            key
        ) in key_change_dict.keys():  # changes stored in the altered key change log wb
            if ws.cell(row=row_num, column=1).value == key:
                ws.cell(row=row_num, column=1).value = key_change_dict[key]
        for year in YEAR_LIST:  # changes to yearly profile keys
            if ws.cell(row=row_num, column=1).value == year + " CDEL Forecast Total":
                ws.cell(row=row_num, column=1).value = (
                    year + " CDEL Forecast one off new costs"
                )

    return wb.save(master_file)


def run_change_keys(master_files_list: list, key_dict: Dict[str, str]) -> None:
    """
    runs code which replaces old key names with new names in master excel workbooks.
    """
    for f in master_files_list:
        alter_wb_master_file_key_names(f, key_dict)


def string_conversion(name):
    if isinstance(name, str):
        return [name]
    else:
        return name


def compare_masters(files: List[typing.TextIO], projects: List[str] or str) -> workbook:
    """Takes two masters and compares if there have been any changes to project data values,
    as well as any new data keys.
    files = list of file paths to masters. Only last two masters are used
    projects = list of those projects that require data checking"""

    projects = string_conversion(projects)

    last_master = project_data_from_master(files[1], 4, 2090)

    wb = load_workbook(files[0])
    ws = wb.active

    project_count = []
    change_count = 0
    new_key_count = 0

    for row_num in range(2, ws.max_row + 1):
        key = ws.cell(row=row_num, column=1).value
        for column_num in range(2, ws.max_column + 1):
            project_name = ws.cell(row=1, column=column_num).value
            if project_name not in projects:
                pass
            else:
                wb_value = ws.cell(row=row_num, column=column_num).value
                try:
                    dict_value = last_master.data[project_name][key]
                    if wb_value == dict_value:
                        pass
                    else:
                        ws.cell(row=row_num, column=column_num).fill = PatternFill(
                            start_color="ffba00", end_color="ffba00", fill_type="solid"
                        )
                        project_count.append(project_name)
                        change_count += 1
                except KeyError:
                    if (
                        project_name in last_master.projects
                    ):  # key error due to key not being present.
                        ws.cell(row=row_num, column=1).fill = PatternFill(
                            start_color="ffba00", end_color="ffba00", fill_type="solid"
                        )
                    else:  # key error due to project not being present.
                        ws.cell(row=1, column=column_num).fill = PatternFill(
                            start_color="ffba00", end_color="ffba00", fill_type="solid"
                        )
    # separate lop to calculate this
    for row_num in range(2, ws.max_row + 1):
        key = ws.cell(row=row_num, column=1).value
        if key not in list(last_master.data[projects[0]].keys()):
            new_key_count += 1

    count_ws = wb.create_sheet("Count", 1)

    count_ws.cell(row=1, column=1).value = "No of changes"
    count_ws.cell(row=1, column=2).value = change_count
    count_ws.cell(row=2, column=1).value = "No of new keys"
    count_ws.cell(row=2, column=2).value = new_key_count

    project_count = Counter(project_count)
    for i, p in enumerate(project_count.keys()):
        count_ws.cell(row=i + 3, column=1).value = p
        count_ws.cell(row=i + 3, column=2).value = project_count[p]

    return wb


def totals_chart(
    fig_size: str,
    costs: CostData,
    benefits: BenefitsData,
    project: str or List[str],
    *args: Tuple[Optional[str]]
) -> None:
    """Small function to hold together code to create and save a total_costs_benefits_bar_chart"""
    costs.get_cost_totals(project, "ipdc_costs")
    benefits.get_ben_totals(project, "ipdc_benefits")
    if args == ():
        f = total_costs_benefits_bar_chart(fig_size, costs, benefits)
        f.savefig(root_path / "output/{}_profile.png".format(costs.entity[0]))
    else:
        f = total_costs_benefits_bar_chart(fig_size, costs, benefits, args[0])
        f.savefig(root_path / "output/{}_profile.png".format(str(args[0])))


def standard_profile(
    fig: plt.figure,
    costs: CostData,
    project: str or List[str],
    *args: Tuple[Optional[str]]
):
    """Small function to hold together code to create and save a cost_profile_graph"""
    costs.get_cost_profile(project, "ipdc_costs")
    if args == ():
        f = cost_profile_graph(fig, costs)
        f.savefig(root_path / "output/{}_profile.png".format(costs.entity[0]))
    else:
        f = cost_profile_graph(fig, costs, args[0])
        f.savefig(root_path / "output/{}_profile.png".format(str(args[0])))


# LIST_OF_GROUPS = [master.current_projects,
#                   Projects.he,
#                   Projects.rail,
#                   Projects.rail_franchising,
#                   Projects.hs2,
#                   Projects.hsmrpg,
#                   Projects.sarh2,
#                   Projects.all_not_hs2,
#                   Projects.fbc_stage,
#                   Projects.obc_stage,
#                   Projects.sobc_stage]
# LIST_OF_TITLES = ['ALL',
#                   'HE',
#                   'RAIL INFRASTRUCTURE',
#                   'RAIL FRANCHISING',
#                   'HS2',
#                   'HSMRPG',
#                   'AMIS (SARH2)',
#                   'ALL, NOT HS2,',
#                   'FBC Projects',
#                   'OBC Projects',
#                   'SOBC Projects']


# def compile_all_profiles():
#     report_doc = open_word_doc(wd_path)
#     for i, p in enumerate(LIST_OF_GROUPS):
#         costs.get_cost_profile(p, 'ipdc_costs')
#         graph = cost_profile_graph(costs, LIST_OF_TITLES[i])
#         put_matplotlib_fig_into_word(report_doc, graph)
#         report_doc.save(root_path / "output/different_cost_profiles.docx")
#
#
# def compile_all_totals():
#     report_doc = open_word_doc(wd_path)
#     for i, p in enumerate(LIST_OF_GROUPS):
#         costs.get_cost_totals(p, 'ipdc_costs')
#         benefits.get_ben_totals(p, 'ipdc_benefits')
#         graph = total_costs_benefits_bar_chart(costs, benefits, LIST_OF_TITLES[i])
#         put_matplotlib_fig_into_word(report_doc, graph)
#         report_doc.save(root_path / "output/different_total_cost_profiles.docx")
