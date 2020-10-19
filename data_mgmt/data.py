import datetime
import difflib
import os
from typing import List, Dict, Union

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import timedelta, date
import numpy as np
from datamaps.api import project_data_from_master
import platform
from pathlib import Path

from docx import Document, table
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, Cm, RGBColor, Inches


def _platform_docs_dir() -> Path:
    #  Cross plaform file path handling
    if platform.system() == "Linux":
        return Path.home() / "Documents" / "analysis_engine"
    if platform.system() == "Darwin":
        return Path.home() / "Documents" / "analysis_engine"
    else:
        return Path.home() / "Documents" / "analysis_engine"


root_path = _platform_docs_dir()


def get_master_data():
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


def get_current_project_names():
    master = project_data_from_master(
        root_path / "core_data/master_2_2020.xlsx", 2, 2020
    )
    return master.projects

def get_project_information():
    return project_data_from_master(
    root_path / "core_data/project_info.xlsx", 1, 2020
    )

# for project summary pages
SRO_conf_table_list = [
    "SRO DCA",
    "Finance DCA",
    "Benefits DCA",
    "Resourcing DCA",
    "Schedule DCA",
]
SRO_conf_key_list = [
    "Departmental DCA",
    "SRO Finance confidence",
    "SRO Benefits RAG",
    "Overall Resource DCA - Now",
    "SRO Schedule Confidence",
]

ipdc_date = datetime.date(
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
    brighton_ml = "Brighton Mainline Upgrade Upgrade Programme (BMUP)"
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
    tru = "Transpennine Route Upgrade (TRU)"
    wrlth = "Western Rail Link to Heathrow"

    # test masters project names
    sot = "Sea of Tranquility"
    a11 = "Apollo 11"
    a13 = "Apollo 13"
    f9 = "Falcon 9"
    columbia = "Columbia"
    mars = "Mars"

    # lists of projects names in groups
    # current_list = master_data_list[0].projects

    he = [lower_thames_crossing, a303, a14, a66, a12, m4, a428, a417, a385]

    hs2 = [hs2_1, hs2_2a, hs2_2b]
    hsmrpg = [
        hs2_1,
        hs2_2a,
        hs2_2b,
        ewr_central,
        ewr_western,
        hexagon,
        northern_powerhouse,
    ]
    ewr = [ewr_config1, ewr_config2, ewr_config3]

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
]

COST_LIST = [" RDEL Forecast Total", " CDEL Forecast Total", " Forecast Non-Gov"]


def current_projects(project_data):
    """Gets list of current/live projects from the project information wb"""
    output_list = []
    for p in project_data.projects:
        if project_data.data[p]["Status"] == "Live":
            output_list.append(p)

    return output_list


class Master:
    def __init__(
        self,
        master_data: List[Dict[str, Union[str, int, date, float]]],
        project_names: list,
    ) -> None:
        self.master_data = master_data
        self.project_names = project_names
        self.bl_info = {}
        self.bl_index = {}
        self.baseline_data()

    def baseline_data(self):

        """
        Returns the two dictionaries baseline_info and baseline_index for all projects for all
        baseline types
        """

        baseline_info = {}
        baseline_index = {}

        for b_type in list(BASELINE_TYPES.keys()):
            project_baseline_info = {}
            project_baseline_index = {}
            for name in self.project_names:
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

    def project_data(self, milestone_type):  # renamed to project_data
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
            for name in self.masters.project_names:
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
            for name in self.masters.project_names:
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
                for name in self.masters.project_names:
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
        self.cat_profile = []
        self.cat_unprofiled = []
        self.total = []
        self.spent = []
        self.profile = []
        self.unprofile = []
        self.current_profile = []
        self.last_profile = []
        self.baseline_profile_one = []
        self.baseline_profile_two = []
        self.rdel_profile = []
        self.cdel_profile = []
        self.ngov_profile = []
        self.current_profile_project = []
        self.last_profile_project = []
        self.baseline_profile_one_project = []
        self.baseline_profile_two_project = []
        self.rdel_profile_project = []
        self.cdel_profile_project = []
        self.ngov_profile_project = []
        # self.cost_totals()
        # self.get_profile()

    def cost_totals(self):
        total = []
        spent = []
        profile = []
        unprofile = []
        cat_spent = []
        cat_profile = []
        cat_unprofiled = []

        for i in reversed(range(3)):  # reversed for matplotlib chart design
            pre_pro_rdel_list = []
            pre_pro_cdel_list = []
            pre_pro_ngov_list = []
            pro_rdel_list = []
            pro_cdel_list = []
            pro_ngov_list = []
            unpro_rdel_list = []
            unpro_cdel_list = []
            unpro_ngov_list = []
            for name in self.master.project_names:
                try:
                    pre_pro_rdel = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Pre-profile RDEL"]
                    pre_pro_cdel = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Pre-profile CDEL"]
                    pre_pro_ngov = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Pre 19-20 Forecast Non-Gov"]
                    if pre_pro_ngov is None:
                        pre_pro_ngov = 0
                    pro_rdel = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Total RDEL Forecast Total"]
                    pro_cdel = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Total CDEL Forecast Total WLC"]
                    pro_ngov = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Non-Gov Total Forecast"]
                    if pro_ngov is None:
                        pro_ngov = 0
                    unpro_rdel = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Unprofiled RDEL Forecast Total"]
                    unpro_cdel = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Unprofiled CDEL Forecast Total WLC"]
                    unpro_ngov = self.master.master_data[
                        self.master.bl_index[name][i]
                    ].data[name]["Unprofiled Forecast Non-Gov"]
                    if unpro_ngov is None:
                        unpro_ngov = 0
                    pre_pro_rdel_list.append(pre_pro_rdel)
                    pre_pro_cdel_list.append(pre_pro_cdel)
                    pre_pro_ngov_list.append(pre_pro_ngov)
                    pro_rdel_list.append(pro_rdel)
                    pro_cdel_list.append(pro_cdel)
                    pro_ngov_list.append(pro_ngov)
                    unpro_rdel_list.append(unpro_rdel)
                    unpro_cdel_list.append(unpro_cdel)
                    unpro_ngov_list.append(unpro_ngov)

                except (TypeError, KeyError):  # KeyError temporary
                    pre_pro_rdel_list.append(0)
                    pre_pro_cdel_list.append(0)
                    pre_pro_ngov_list.append(0)
                    pro_rdel_list.append(0)
                    pro_cdel_list.append(0)
                    pro_ngov_list.append(0)
                    unpro_rdel_list.append(0)
                    unpro_cdel_list.append(0)
                    unpro_ngov_list.append(0)

            total_rdel_pre_pro = sum(pre_pro_rdel_list)
            total_cdel_pre_pro = sum(pre_pro_cdel_list)
            total_ngov_pre_pro = sum(pre_pro_ngov_list)
            total_rdel_pro = sum(pro_rdel_list)
            total_cdel_pro = sum(pro_cdel_list)
            total_ngov_pro = sum(pro_ngov_list)
            total_rdel_unpro = sum(unpro_rdel_list)
            total_cdel_unpro = sum(unpro_cdel_list)
            total_ngov_unpro = sum(unpro_ngov_list)

            if i == 0:
                cat_spent.append(total_rdel_pre_pro)
                cat_spent.append(total_cdel_pre_pro)
                cat_spent.append(total_ngov_pre_pro)
                rdel_pro = total_rdel_pro - (total_rdel_pre_pro + total_rdel_unpro)
                cdel_pro = total_cdel_pro - (total_cdel_pre_pro + total_cdel_unpro)
                ngov_pro = total_ngov_pro - (total_ngov_pre_pro + total_ngov_unpro)
                cat_profile.append(rdel_pro)
                cat_profile.append(cdel_pro)
                cat_profile.append(ngov_pro)
                cat_unprofiled.append(total_rdel_unpro)
                cat_unprofiled.append(total_cdel_unpro)
                cat_unprofiled.append(total_ngov_unpro)

            total_pre_pro = (
                sum(pre_pro_rdel_list) + sum(pre_pro_cdel_list) + sum(pre_pro_ngov_list)
            )
            total_unpro = (
                sum(unpro_rdel_list) + sum(unpro_cdel_list) + sum(unpro_ngov_list)
            )
            total_pro = (
                sum(pro_rdel_list) + sum(pro_cdel_list) + sum(pro_ngov_list)
            ) - (total_pre_pro + total_unpro)
            total.append((sum(pro_rdel_list) + sum(pro_cdel_list)) + sum(pro_ngov_list))
            spent.append(total_pre_pro)
            profile.append(total_pro)
            unprofile.append(total_unpro)

        self.cat_spent = cat_spent
        self.cat_profile = cat_profile
        self.cat_unprofiled = cat_unprofiled
        self.total = total
        self.spent = spent
        self.profile = profile
        self.unprofile = unprofile

    def get_profile_all_old(self):
        year_list = [
            "20-21",
            "21-22",
            "22-23",
            "23-24",
            "24-25",
            "25-26",
            "26-27",
            "27-28",
            "28-29",
        ]

        cost_list = [
            " RDEL Forecast Total",
            " CDEL Forecast Total",
            " Forecast Non-Gov",
        ]

        current_profile = []
        last_profile = []
        baseline_profile = []
        for i in range(3):
            profile = []
            for year in year_list:
                a_list = []
                for cost_type in cost_list:
                    data = []
                    for name in self.master.project_names:
                        try:
                            cost = self.master.master_data[
                                self.master.bl_index[name][i]
                            ].data[name][year + cost_type]
                            if cost is None:
                                cost = 0
                        except (KeyError, TypeError):  # to handle baselines
                            cost = 0
                        data.append(cost)
                    a_list.append(sum(data))
                profile.append(sum(a_list))

            if i == 0:
                current_profile = profile
            if i == 1:
                last_profile = profile
            if i == 2:
                baseline_profile = profile

        self.current_profile = current_profile
        self.last_profile = last_profile
        self.baseline_profile = baseline_profile

    def get_profile_all(self, baseline: str) -> None:
        """Returns several lists which contain the sum of different cost profiles for the group of project
        contained with the master"""
        self.baseline = baseline

        current_profile = []
        last_profile = []
        baseline_profile_one = []
        baseline_profile_two = []
        rdel_current_profile = []
        cdel_current_profile = []
        ngov_current_profile = []

        missing_projects = []

        for i in range(3):
            yearly_profile = []
            rdel_yearly_profile = []
            cdel_yearly_profile = []
            ngov_yearly_profile = []
            for year in YEAR_LIST:
                cost_total = 0
                rdel_total = 0
                cdel_total = 0
                ngov_total = 0
                for cost_type in COST_LIST:
                    for project in self.master.project_names:
                        project_bl_index = self.master.bl_index[baseline][project]
                        try:
                            cost = self.master.master_data[project_bl_index[i]].data[
                                project
                            ][year + cost_type]
                            if cost is None:
                                cost = 0
                            cost_total += cost
                        except KeyError:  # to handle like for likeness comparison between current and last quarter.
                            # work also required on the data set so that the year key doesn't throw a key error due
                            # to all years not be present in all masters. This is why there as a messy extra step below.
                            p_master_data_keys = self.master.master_data[
                                project_bl_index[i]
                            ].data[project]
                            concatenated_key = year + cost_type
                            if concatenated_key in p_master_data_keys:
                                print(
                                    "NOTE: "
                                    + str(project)
                                    + " was not part of the portfolio last quarter. This "
                                    "means current quarter and last quarter cost profiles cannot be compared as like"
                                    "for like unless "
                                    + str(project)
                                    + " is removed from this group."
                                )
                            cost = 0
                            cost_total += cost
                        except IndexError:  # to handle projects lacking baseline indexes. This requires changes to the data.
                            print(
                                str(project)
                                + " has no "
                                + str(self.baseline)
                                + " baseline. All projects must have at least"
                                "one baseline point. Even if this is only the point at which it entered the portfolio"
                                ". Therefore this programme is stopping until a baseline index is provided"
                            )
                            #  The programme should stop here but for some reason isn't.
                        except TypeError:  # Handles projects not present in the previous quarter
                            missing_projects.append(str(project))  # projects added here. message is below.
                            cost = 0
                            cost_total += cost

                        if cost_type == " RDEL Forecast Total":
                            rdel_total += cost
                        if cost_type == " CDEL Forecast Total":
                            cdel_total += cost
                        if cost_type == " Forecast Non-Gov":
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

        missing_projects = list(set(missing_projects))  # if TypeError raised above
        if len(missing_projects) is not 0:
            print(
                "NOTE: "
                + str(missing_projects)
                + " was not part of the portfolio last quarter. This "
                  "means current quarter and last quarter cost profiles cannot be compared as like"
                  "for like unless they are removed from this group."
            )

        self.current_profile = current_profile
        self.last_profile = last_profile
        self.baseline_profile_one = baseline_profile_one
        self.baseline_profile_two = baseline_profile_two
        self.rdel_profile = rdel_current_profile
        self.cdel_profile = cdel_current_profile
        self.ngov_profile = ngov_current_profile

    def get_profile_project(self, project_name: str, baseline: str) -> list:
        """Returns several lists which contain different cost profiles for a given project"""
        self.project_name = project_name
        self.baseline = baseline

        current_profile = []
        last_profile = []
        baseline_profile_one = []
        baseline_profile_two = []
        rdel_current_profile = []
        cdel_current_profile = []
        ngov_current_profile = []

        cost_bl_index = self.master.bl_index[baseline][self.project_name]
        for i in range(len(cost_bl_index)):
            yearly_profile = []
            rdel_yearly_profile = []
            cdel_yearly_profile = []
            ngov_yearly_profile = []
            for year in YEAR_LIST:
                cost_total = 0
                for cost_type in COST_LIST:
                    try:
                        cost = self.master.master_data[cost_bl_index[i]].data[
                            self.project_name
                        ][year + cost_type]
                        if cost is None:
                            cost = 0
                        cost_total += cost
                    except KeyError:  # to handle data across different financial years
                        cost = 0
                        cost_total += cost
                    if cost_type == " RDEL Forecast Total":
                        rdel_total = cost
                    if cost_type == " CDEL Forecast Total":
                        cdel_total = cost
                    if cost_type == " Forecast Non-Gov":
                        ngov_total = cost

                yearly_profile.append(cost_total)
                rdel_yearly_profile.append(rdel_total)
                cdel_yearly_profile.append(cdel_total)
                ngov_yearly_profile.append(ngov_total)

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

        self.current_profile_project = current_profile
        self.last_profile_project = last_profile
        self.baseline_profile_one_project = baseline_profile_one
        self.baseline_profile_two_project = baseline_profile_two
        self.rdel_profile_project = rdel_current_profile
        self.cdel_profile_project = cdel_current_profile
        self.ngov_profile_project = ngov_current_profile


class BenefitsData:
    def __init__(self, masters_object):
        self.masters = masters_object
        self.total = []
        self.achieved = []
        self.profile = []
        self.unprofile = []
        self.cat_achieved = []
        self.cat_profile = []
        self.cat_unprofile = []
        self.disbenefit = []
        self.ben_totals()

    def ben_totals(self):
        """given a list of project names returns benefit
        data lists for placement in matplotlib charts
        """

        ben_key_list = [
            "Pre-profile BEN Total",
            "Total BEN Forecast - Total Monetised Benefits",
            "Unprofiled Remainder BEN Forecast - Total Monetised Benefits",
        ]

        ben_type_key_list = [
            (
                "Pre-profile BEN Forecast Gov Cashable",
                "Pre-profile BEN Forecast Gov Non-Cashable",
                "Pre-profile BEN Forecast - Economic (inc Private Partner)",
                "Pre-profile BEN Forecast - Disbenefit UK Economic",
            ),
            (
                "Unprofiled Remainder BEN Forecast - Gov. Cashable",
                "Unprofiled Remainder BEN Forecast - Gov. Non-Cashable",
                "Unprofiled Remainder BEN Forecast - Economic (inc Private Partner)",
                "Unprofiled Remainder BEN Forecast - Disbenefit UK Economic",
            ),
            (
                "Total BEN Forecast - Gov. Cashable",
                "Total BEN Forecast - Gov. Non-Cashable",
                "Total BEN Forecast - Economic (inc Private Partner)",
                "Total BEN Forecast - Disbenefit UK Economic",
            ),
        ]

        total = []
        achieved = []
        profile = []
        unprofile = []

        for i in reversed(range(3)):
            ben_achieved = []
            ben_profile = []
            ben_unprofile = []
            for y in ben_key_list:
                for name in self.masters.project_names:
                    try:
                        ben = self.masters.master_data[
                            self.masters.bl_index[name][i]
                        ].data[name][y]
                    except TypeError:
                        ben = 0

                    if y is ben_key_list[0]:
                        ben_achieved.append(ben)
                    if y is ben_key_list[1]:
                        ben_profile.append(ben)
                    if y is ben_key_list[2]:
                        ben_unprofile.append(ben)

            achieved.append(sum(ben_achieved))
            profile.append(sum(ben_profile) - (sum(ben_achieved) + sum(ben_unprofile)))
            unprofile.append(sum(ben_unprofile))
            total.append(sum(ben_profile))

        cat_achieved = []
        cat_profile = []
        cat_unprofile = []
        disbenefit = []

        for x in range(4):
            ben_cat_achieved = []
            ben_cat_profile = []
            ben_cat_unprofile = []
            for y in ben_type_key_list:
                for name in self.masters.project_names:

                    ben = self.masters.master_data[0].data[name][y[x]]
                    if ben is None:
                        ben = 0

                    if y is ben_type_key_list[0]:
                        ben_cat_achieved.append(ben)
                    if y is ben_type_key_list[1]:
                        ben_cat_profile.append(ben)
                    if y is ben_type_key_list[2]:
                        ben_cat_unprofile.append(ben)

                    if "Disbenefit" in y[x]:
                        disbenefit.append(ben)

            cat_achieved.append(sum(ben_cat_achieved))
            cat_profile.append(
                sum(ben_cat_profile) - (sum(ben_cat_achieved) + sum(ben_cat_unprofile))
            )
            cat_unprofile.append(sum(ben_cat_unprofile))
            disbenefit.append(sum(disbenefit))

        self.total = total
        self.achieved = achieved
        self.profile = profile
        self.unprofile = unprofile
        self.cat_achieved = cat_achieved
        self.cat_profile = cat_profile
        self.cat_unprofile = cat_unprofile
        self.disbenefit = disbenefit


# for compiling vfm data in matplotlib chart
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


def project_cost_profile_graph(cost_master: object):
    """Compiles a matplotlib line chart for cost profile contained within cost_master class.
    It creates two plots. First plot shows overall profile in current, last quarters anb
    baseline form. Second plot shows rdel, cdel, and 'non-gov' cost profile"""

    fig, (ax1, ax2) = plt.subplots(2)  # two subplots for this chart

    """cost profile charts"""
    fig.suptitle(
        str(cost_master.project_name) + " Cost Profile", fontweight="bold"
    )  # title

    # Overall cost profile chart
    try:  # try statement handles the project having no baseline profile.
        ax1.plot(
            YEAR_LIST,
            np.array(cost_master.baseline_profile_one_project),
            label="Baseline",
            linewidth=3.0,
            marker="o",
        )
    except ValueError:
        pass
    ax1.plot(
        YEAR_LIST,
        np.array(cost_master.last_profile_project),
        label="Last quarter",
        linewidth=3.0,
        marker="o",
    )
    ax1.plot(
        YEAR_LIST,
        np.array(cost_master.current_profile_project),
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

    # plot rdel/cdel chart data
    if (
        sum(cost_master.ngov_profile_project) != 0
    ):  # if statement as most projects don't have ngov cost.
        ax2.plot(
            YEAR_LIST,
            np.array(cost_master.ngov_profile_project),
            label="Non-Gov",
            linewidth=3.0,
            marker="o",
        )
    ax2.plot(
        YEAR_LIST,
        np.array(cost_master.cdel_profile_project),
        label="CDEL",
        linewidth=3.0,
        marker="o",
    )
    ax2.plot(
        YEAR_LIST,
        np.array(cost_master.rdel_profile_project),
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

    #plt.show()

    return fig


def group_cost_profile_graph(cost_master: object, title: str):
    """Compiles a matplotlib line chart for costs of all projects contained within cost_master class.
    As as default last quarters profile is not included. It creates two plots. First plot shows overall
    profile in current, last quarters anb baseline form. Second plot shows rdel, cdel, and 'non-gov' cost profile"""

    fig, (ax1, ax2) = plt.subplots(2)  # two subplots for this chart

    """cost profile charts"""
    fig.suptitle(title + " Cost Profile", fontweight="bold")  # title

    # Overall cost profile chart
    ax1.plot(
        YEAR_LIST,
        np.array(cost_master.baseline_profile_one),  # baseline profile
        label="Baseline",
        linewidth=3.0,
        marker="o",
    )
    ax1.plot(
        YEAR_LIST,
        np.array(cost_master.last_profile),  # last quarter profile
        label='Last quarter',
        linewidth=3.0,
        marker="o"
    )
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

    # plot rdel/cdel chart data
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


def spent_calculation(master: Dict[str, Union[str, date, int, float]], project: str) -> int:
    keys = ['Pre-profile RDEL', '20-21 RDEL STD Total', 'Pre-profile CDEL', '20-21 CDEL STD Total']

    total = 0
    for k in keys:
        total += master.data[project][k]

    return total


def open_word_doc(wd_path: str) -> Document:
    """Function stores an empty word doc as a variable"""
    return Document(wd_path)


def wd_heading(doc: Document, project_info: Dict[str, str], project_name: str) -> None:
    """Function adds header to word doc"""
    font = doc.styles['Normal'].font
    font.name = 'Arial'
    font.size = Pt(12)

    heading = str(project_info.data[project_name]['Abbreviations'])
    intro = doc.add_heading(str(heading), 0)
    intro.alignment = 1
    intro.bold = True


def key_contacts(doc: Document, master: Master, project_name: str) -> None:
    """Function adds key contact details"""
    sro_name = master.master_data[0].data[project_name]['Senior Responsible Owner (SRO)']
    if sro_name is None:
        sro_name = 'tbc'

    sro_email = master.master_data[0].data[project_name]['Senior Responsible Owner (SRO) - Email']
    if sro_email is None:
        sro_email = 'email: tbc'

    sro_phone = master.master_data[0].data[project_name]['SRO Phone No.']
    if sro_phone == None:
        sro_phone = 'phone number: tbc'

    doc.add_paragraph('SRO: ' + str(sro_name) + ', ' + str(sro_email) + ', ' + str(sro_phone))

    pd_name = master.master_data[0].data[project_name]['Project Director (PD)']
    if pd_name is None:
        pd_name = 'TBC'

    pd_email = master.master_data[0].data[project_name]['Project Director (PD) - Email']
    if pd_email is None:
        pd_email = 'email: tbc'

    pd_phone = master.master_data[0].data[project_name]['PD Phone No.']
    if pd_phone is None:
        pd_phone = 'TBC'

    doc.add_paragraph('PD: ' + str(pd_name) + ', ' + str(pd_email) + ', ' + str(pd_phone))

    contact_name = master.master_data[0].data[project_name]['Working Contact Name']
    if contact_name is None:
        contact_name = 'TBC'

    contact_email = master.master_data[0].data[project_name]['Working Contact Email']
    if contact_email is None:
        contact_email = 'email: tbc'

    contact_phone = master.master_data[0].data[project_name]['Working Contact Telephone']
    if contact_phone is None:
        contact_phone = 'TBC'

    doc.add_paragraph('PfM reporting lead: ' + str(contact_name) + ', ' + str(contact_email)
                      + ', ' + str(contact_phone))


def dca_table(doc: Document, master: Master, project_name: str) -> None:
    """Creates SRO confidence table"""
    w_table = doc.add_table(rows=1, cols=5)
    hdr_cells = w_table.rows[0].cells
    hdr_cells[0].text = 'Delivery confidence'
    hdr_cells[1].text = 'This quarter'
    hdr_cells[2].text = str(master.master_data[1].quarter)
    hdr_cells[3].text = str(master.master_data[2].quarter)
    hdr_cells[4].text = str(master.master_data[3].quarter)

    for x, dca_key in enumerate(SRO_conf_key_list):
        row_cells = w_table.add_row().cells
        row_cells[0].text = dca_key
        for i, m in enumerate(master.master_data[:4]):  # last four masters taken
            try:
                rating = convert_rag_text(m.data[project_name][dca_key])
                row_cells[i + 1].text = rating
                cell_colouring(row_cells[i + 1], rating)
            except (KeyError, TypeError):
                row_cells[i + 1].text = "N/A"

    w_table.style = 'Table Grid'
    make_rows_bold([w_table.rows[0]])  # makes top of table bold.
    # make_columns_bold([table.columns[0]]) #right cells in table bold
    column_widths = (Cm(3.9), Cm(2.9), Cm(2.9), Cm(2.9), Cm(2.9))
    set_col_widths(w_table, column_widths)


def dca_narratives(doc: Document, master: Master, project_name: str) -> None:
    """Places all narratives into document and checks for differences between
    current and last quarter"""

    doc.add_paragraph()
    p = doc.add_paragraph()
    text = '*Red text highlights changes in narratives from last quarter'
    p.add_run(text).font.color.rgb = RGBColor(255, 0, 0)

    headings_list = ['SRO delivery confidence narrative',
                     'Financial cost narrative',
                     'Financial comparison with last quarter',
                     'Financial comparison with baseline',
                     'Benefits Narrative',
                     'Benefits comparison with last quarter',
                     'Benefits comparison with baseline',
                     'Milestone narrative']

    narrative_keys_list = ['Departmental DCA Narrative',
                           'Project Costs Narrative',
                           'Cost comparison with last quarters cost narrative',
                           'Cost comparison within this quarters cost narrative',
                           'Benefits Narrative',
                           'Ben comparison with last quarters cost - narrative',
                           'Ben comparison within this quarters cost - narrative',
                           'Milestone Commentary']

    for x in range(len(headings_list)):
        doc.add_paragraph().add_run(str(headings_list[x])).bold = True
        text_one = str(master.master_data[0].data[project_name][narrative_keys_list[x]])
        try:
            text_two = str(master.master_data[1].data[project_name][narrative_keys_list[x]])
        except KeyError:
            text_two = text_one

        # There are two options here for comparing text. Have left this for now.
        # compare_text_showall(dca_a, dca_b, doc)
        compare_text_new_and_old(text_one, text_two, doc)


def year_cost_profile_chart(doc: Document, cost_master: CostData) -> None:
    """Places line graph cost profile into word document"""

    new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)  # new page
    # change to landscape
    new_width, new_height = new_section.page_height, new_section.page_width
    new_section.orientation = WD_ORIENTATION.LANDSCAPE
    new_section.page_width = new_width
    new_section.page_height = new_height

    fig = project_cost_profile_graph(cost_master)

    # Size and shape of figure.
    fig.canvas.draw()
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

    # Place fig in word doc.
    fig.show()
    fig.savefig('cost_profile.png')
    doc.add_picture('cost_profile.png', width=Inches(8))  # to place nicely in doc
    os.remove('cost_profile.png')


def convert_rag_text(dca_rating: str) -> None:
    """Converts RAG name into a acronym"""

    if dca_rating == 'Green':
        return 'G'
    elif dca_rating == 'Amber/Green':
        return 'A/G'
    elif dca_rating == 'Amber':
        return 'A'
    elif dca_rating == 'Amber/Red':
        return 'A/R'
    elif dca_rating == 'Red':
        return 'R'
    else:
        return ''


def cell_colouring(word_table_cell: table.Table.cell, colour: str) -> None:
    """Function that handles cell colouring for word documents"""

    try:
        if colour == 'R':
            colour = parse_xml(r'<w:shd {} w:fill="cb1f00"/>'.format(nsdecls('w')))
        elif colour == 'A/R':
            colour = parse_xml(r'<w:shd {} w:fill="f97b31"/>'.format(nsdecls('w')))
        elif colour == 'A':
            colour = parse_xml(r'<w:shd {} w:fill="fce553"/>'.format(nsdecls('w')))
        elif colour == 'A/G':
            colour = parse_xml(r'<w:shd {} w:fill="a5b700"/>'.format(nsdecls('w')))
        elif colour == 'G':
            colour = parse_xml(r'<w:shd {} w:fill="17960c"/>'.format(nsdecls('w')))

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

        if diff[i][0:3] == '  |':
            j = i + 1
            if diff[i][0:3] and diff[a][0:3] == '  |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '+ |':
            if diff[i][0:3] and diff[a][0:3] == '+ |':
                y = doc.add_paragraph()
            else:
                pass
        elif diff[i][0:3] == '- |':
            pass
        elif diff[i][0:3] == '  -':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0:3] == '  •':
            y = doc.add_paragraph()
            g = diff[i][2]
            y.add_run(g)
        elif diff[i][0] == '+':
            w = len(diff[i])
            g = diff[i][1:w]
            y.add_run(g).font.color.rgb = RGBColor(255, 0, 0)
        elif diff[i][0] == '-':
            pass
        elif diff[i][0] == '?':
            pass
        else:
            if diff[i] != '+ |':
                y.add_run(diff[i][1:])


