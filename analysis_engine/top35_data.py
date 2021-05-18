import datetime
import math
from collections import Counter
from typing import List, Dict, Union
from datetime import date
from analysis_engine.data import (
    get_group,
    open_word_doc,
    make_file_friendly,
    compare_text_new_and_old,
    set_col_widths,
    make_columns_bold,
    get_iter_list,
    get_milestone_date, get_correct_p_data, COLOUR_DICT, dandelion_number_text, cal_group_angle, convert_rag_text,
    Master, logger, MilestoneData,
)
from datamaps.api import project_data_from_master
import platform
from pathlib import Path
from docx import Document
from dateutil import parser
from dateutil.parser import ParserError

from docx.enum.section import WD_SECTION_START
from docx.shared import Cm, RGBColor, Pt


def _top_35_platform_docs_dir() -> Path:
    #  Cross plaform file path handling
    if platform.system() == "Linux":
        return Path.home() / "Documents" / "top_250"
    if platform.system() == "Darwin":
        return Path.home() / "Documents" / "top_250"
    else:
        return Path.home() / "Documents" / "top_250"


top35_root_path = _top_35_platform_docs_dir()


def top35_get_master_data() -> List[
    Dict[str, Union[str, int, datetime.date, float]]
]:  # how specify a list of dictionaries?
    """Returns a list of dictionaries each containing quarter data"""
    master_data_list = [
        project_data_from_master(
            top35_root_path / "core_data/250_master_april_21.xlsx", 4, 2020
        ),
        project_data_from_master(
            top35_root_path / "core_data/250_master_april_21.xlsx", 3, 2020
        ),
    ]
    return master_data_list


def top35_get_project_information() -> Dict[str, Union[str, int]]:
    """Returns dictionary containing all project meta data"""
    return project_data_from_master(
        top35_root_path / "core_data/top_250_project_info.xlsx", 2, 2020
    )


# class Master:
#     def __init__(
#         self,
#         master_data: List[Dict[str, Union[str, int, datetime.date, float]]],
#         project_information: Dict[str, Union[str, int]],
#     ) -> None:
#         self.master_data = master_data
#         self.project_information = project_information
#         self.current_quarter = self.master_data[0].quarter
#         self.current_projects = self.master_data[0].projects
#         self.abbreviations = {}
#         self.full_names = {}
#         self.bl_info = {}
#         self.bl_index = {}
#         self.dft_groups = {}
#         self.project_group = {}
#         self.project_stage = {}
#         self.quarter_list = []
#         self.get_quarter_list()
#         # self.get_baseline_data()
#         self.check_project_information()
#         self.get_project_abbreviations()
#         # self.check_baselines()
#         self.get_project_groups()
#
#     """This is the entry point for all data. It converts a list of excel wbs (note at the moment)
#     this is actually done prior to being passed into the Master class. The Master class does a number
#     of things.
#     compiles and checks all baseline data for projects. These index reference points.
#     compiles lists of different project groups. e.g stage and DfT Group
#     gets a list of projects currently in the portfolio.
#     checks data returned by projects is consistent with whats in project_information
#     gets project abbreviations
#
#     """
#
#     def get_project_abbreviations(self) -> None:
#         """gets the abbreviations for all current projects.
#         held in the project info document"""
#         abb_dict = {}
#         fn_dict = {}
#         error_case = []
#         for p in self.project_information.projects:
#             abb = self.project_information[p]["Abbreviations"]
#             abb_dict[p] = {"abb": abb, "full name": p}
#             fn_dict[abb] = p
#             if abb is None:
#                 error_case.append(p)
#
#         if error_case:
#             for p in error_case:
#                 logger.critical("No abbreviation provided for " + p + ".")
#             raise ProjectNameError(
#                 "Abbreviations must be provided for all projects in project_info. Program stopping. Please amend"
#             )
#
#         self.abbreviations = abb_dict
#         self.full_names = fn_dict
#
#     # def get_baseline_data(self) -> None:
#     #     """
#     #     Returns the two dictionaries baseline_info and baseline_index for all projects for all
#     #     baseline types
#     #     """
#     #
#     #     baseline_info = {}
#     #     baseline_index = {}
#     #
#     #     for b_type in list(BASELINE_TYPES.keys()):
#     #         project_baseline_info = {}
#     #         project_baseline_index = {}
#     #         for name in self.current_projects:
#     #             bc_list = []
#     #             lower_list = []
#     #             for i, master in reversed(list(enumerate(self.master_data))):
#     #                 if name in master.projects:
#     #                     try:
#     #                         approved_bc = master.data[name][b_type]
#     #                         quarter = str(master.quarter)
#     #                     # exception handling in here in case data keys across masters are not consistent.
#     #                     # not sure this is necessary any more
#     #                     except KeyError:
#     #                         print(
#     #                             str(b_type)
#     #                             + " keys not present in "
#     #                             + str(master.quarter)
#     #                         )
#     #                     if approved_bc == "Yes":
#     #                         bc_list.append(approved_bc)
#     #                         lower_list.append((approved_bc, quarter, i))
#     #                 else:
#     #                     pass
#     #             for i in reversed(range(2)):
#     #                 if name in self.master_data[i].projects:
#     #                     approved_bc = self.master_data[i][name][b_type]
#     #                     quarter = str(self.master_data[i].quarter)
#     #                     lower_list.append((approved_bc, quarter, i))
#     #                 else:
#     #                     quarter = str(self.master_data[i].quarter)
#     #                     lower_list.append((None, quarter, None))
#     #
#     #             index_list = []
#     #             for x in lower_list:
#     #                 index_list.append(x[2])
#     #
#     #             project_baseline_info[name] = list(reversed(lower_list))
#     #             project_baseline_index[name] = list(reversed(index_list))
#     #
#     #         baseline_info[BASELINE_TYPES[b_type]] = project_baseline_info
#     #         baseline_index[BASELINE_TYPES[b_type]] = project_baseline_index
#     #
#     #     self.bl_info = baseline_info
#     #     self.bl_index = baseline_index
#
#     def check_project_information(self) -> None:
#         """Checks that project names in master are present/the same as in project info.
#         Stops the programme if not"""
#         error_cases = []
#         for p in self.current_projects:
#             if p not in self.project_information.projects:
#                 error_cases.append(p)
#
#         if error_cases:
#             for p in error_cases:
#                 logger.critical(p + " has not been found in the project_info document.")
#             raise ProjectNameError(
#                 "Project names in "
#                 + str(self.master_data[0].quarter)
#                 + " master and project_info must match. Program stopping. Please amend."
#             )
#         else:
#             logger.info("The latest master and project information match")
#
#     # def check_baselines(self) -> None:
#     #     """checks that projects have the correct baseline information. stops the
#     #     programme if baselines are missing"""
#     #     for v in IPDC_BASELINE_TYPES.values():
#     #         for p in self.current_projects:
#     #             baselines = self.bl_index[v][p]
#     #             if len(baselines) <= 2:
#     #                 logger.critical(
#     #                     p
#     #                     + " does not have a baseline point for "
#     #                     + v
#     #                     + " this could cause the programme to "
#     #                       "crash. Therefore the programme is stopping. "
#     #                       "Please amend the data for " + p + " so that "
#     #                                                          " it has at least one baseline point for " + v
#     #                 )
#
#     def get_project_groups(self) -> None:
#         """gets the groups that projects are part of e.g. business case
#         stage or dft group"""
#
#         raw_dict = {}
#         raw_list = []
#         group_list = []
#         stage_list = []
#         for i, master in enumerate(self.master_data):
#             lower_dict = {}
#             for p in master.projects:
#                 dft_group = self.project_information[p][
#                     "Group"
#                 ]  # different groups cleaned here
#                 if dft_group is None:
#                     logger.critical(
#                         str(p)
#                         + " does not have a Group value in the project information document."
#                     )
#                     raise ProjectGroupError(
#                         "Program stopping as this could cause a crash. Please check project Group info."
#                     )
#                 if dft_group not in list(DFT_GROUP_DICT.keys()):
#                     logger.critical(
#                         str(p)
#                         + " Group value is "
#                         + str(dft_group)
#                         + " . This is not a recognised group"
#                     )
#                     raise ProjectGroupError(
#                         "Program stopping as this could cause a crash. Please check project Group info."
#                     )
#                 # stage = BC_STAGE_DICT[master[p]["IPDC approval point"]]
#                 raw_list.append(("group", dft_group))
#                 # raw_list.append(("stage", stage))
#                 lower_dict[p] = dict(raw_list)
#                 group_list.append(dft_group)
#                 # stage_list.append(stage)
#
#             raw_dict[str(master.quarter)] = lower_dict
#
#         group_list = list(set(group_list))
#         # stage_list = list(set(stage_list))
#
#         group_dict = {}
#         for i, quarter in enumerate(raw_dict.keys()):
#             lower_g_dict = {}
#             for group_type in group_list:
#                 g_list = []
#                 for p in raw_dict[quarter].keys():
#                     p_group = raw_dict[quarter][p]["group"]
#                     if p_group == group_type:
#                         g_list.append(p)
#                 lower_g_dict[group_type] = g_list
#
#             gmpp_list = []
#             for p in self.master_data[i].projects:
#                 gmpp = self.project_information[p]["GMPP"]
#                 if gmpp is not None:
#                     gmpp_list.append(p)
#                 lower_g_dict["GMPP"] = gmpp_list
#
#             group_dict[quarter] = lower_g_dict
#
#         stage_dict = {}
#         for quarter in raw_dict.keys():
#             lower_s_dict = {}
#             for stage_type in stage_list:
#                 s_list = []
#                 for p in raw_dict[quarter].keys():
#                     p_stage = raw_dict[quarter][p]["stage"]
#                     if p_stage == stage_type:
#                         s_list.append(p)
#                 if stage_type is None:
#                     if s_list:
#                         if quarter == self.current_quarter:
#                             for x in s_list:
#                                 logger.critical(str(x) + " has no IPDC stage date")
#                                 raise ProjectStageError(
#                                     "Programme stopping as this could cause incomplete analysis"
#                                 )
#                         else:
#                             for x in s_list:
#                                 logger.warning(
#                                     "In "
#                                     + str(quarter)
#                                     + " master "
#                                     + str(x)
#                                     + " IPDC stage data is currently None. Please amend."
#                                 )
#                 lower_s_dict[stage_type] = s_list
#             stage_dict[quarter] = lower_s_dict
#
#         self.dft_groups = group_dict
#         self.project_stage = stage_dict
#
#     def get_quarter_list(self) -> None:
#         output_list = []
#         for master in self.master_data:
#             output_list.append(str(master.quarter))
#         self.quarter_list = output_list


def top35_run_p_reports(master: Master, **kwargs) -> None:
    group = get_group(master, str(master.current_quarter), kwargs)
    for p in group:
        p_name = master.project_information.data[p]["Abbreviations"]
        print("Compiling summary for " + p_name)
        report_doc = open_word_doc(top35_root_path / "input/summary_temp.docx")
        output = compile_p_report(report_doc, master, p, **kwargs)
        output.save(
            top35_root_path / "output/{}_report.docx".format(p_name)
        )  # add quarter here


# def run_p_reports_single(master: Master, **kwargs) -> None:
#     group = get_group(master, str(master.current_quarter), kwargs)
#
#     report_doc = open_word_doc(top35_root_path / "input/summary_temp.docx")
#     for p in group:
#         print("Compiling summary for " + p)
#         qrt = make_file_friendly(str(master.master_data[0].quarter))
#         output = compile_p_report(report_doc, get_project_information(), master, p)
#         report_doc.add_section(WD_SECTION_START.NEW_PAGE)
#
#     output.save(
#         top35_root_path / "output/250_report.docx".format(p, qrt)
#     )  # add quarter here


def run_pm_one_lines_single(master: Master, prog_info: Dict, **kwargs) -> None:
    group = get_group(master, str(master.current_quarter), kwargs)

    name_list = []
    for p in group:
        name_list.append((p, prog_info.data[p]["ID Number"]))

    name_list.sort(key=lambda x: x[1])

    report_doc = open_word_doc(top35_root_path / "input/summary_temp.docx")
    for p in name_list:
        p_master = master.master_data[0].data[p[0]]
        p_name = prog_info.data[p[0]]["ID Number"]
        output = deliverables(report_doc, p_master, p_name=p_name)
        # output = pm_one_line(report_doc, p_master, prog_info.data[p[0]]["ID Number"])

    output.save(top35_root_path / "output/deliverables.docx")  # add quarter here


def compile_p_report(
    doc: Document,
    master: Master,
    project_name: str,
    **kwargs,
) -> Document:
    p_master = master.master_data[0].data[project_name]
    r_args = [doc, p_master]
    wd_heading(doc, master.project_information, project_name)
    key_contacts(*r_args)
    project_scope_text(*r_args)
    deliverables(*r_args)
    project_report_meta_data(*r_args)
    # doc.add_section(WD_SECTION_START.NEW_PAGE)
    dca_narratives(*r_args)
    kwargs["group"] = [project_name]
    ms = MilestoneData(master, "ipdc_milestones", **kwargs)  # milestones
    print_out_project_milestones(doc, ms)
    cs = CentralSupportData(master, group=[project_name], quarter=["Q4 20/21"]) # central support
    print_out_central_support(doc, cs)
    return doc


def wd_heading(
        doc: Document, project_info: Dict[str, Union[str, int]], project_name: str
) -> None:
    """Function adds header to word doc"""
    font = doc.styles["Normal"].font
    font.name = "Arial"
    font.size = Pt(12)

    heading = str(
        project_info.data[project_name]["ID Number"]
    )  # integrate into master
    intro = doc.add_heading(str(heading), 0)
    intro.alignment = 1
    intro.bold = True


def key_contacts(doc: Document, p_master: Dict) -> None:
    """Function adds keys contact details"""
    sro_name = p_master["SRO NAME"]
    if sro_name is None:
        sro_name = "tbc"

    sro_email = p_master["SRO EMAIL"]
    if sro_email is None:
        sro_email = "email: tbc"

    # sro_phone = master.master_data[0].data[project_name]["SRO Phone No."]
    # if sro_phone == None:
    #     sro_phone = "phone number: tbc"

    # doc.add_paragraph(
    #     "SRO: " + str(sro_name) + ", " + str(sro_email) + ", " + str(sro_phone)
    # )
    doc.add_paragraph("SRO: " + str(sro_name) + ", " + str(sro_email))


def project_scope_text(doc: Document, p_master: Master) -> Document:
    doc.add_paragraph().add_run("Short project description").bold = True
    text_one = str(p_master["SHORT PROJECT DESCRIPTION"])
    try:
        text_two = str(p_master["SHORT PROJECT DESCRIPTION"])
    except IndexError:
        text_two = text_one
    compare_text_new_and_old(text_one, text_two, doc)
    return doc


def deliverables(doc: Document, p_master: Dict, **kwargs) -> Document:
    dels = [
        p_master["TOP 3 PROJECT DELIVERABLES 1"],  # deliverables
        p_master["TOP 3 PROJECT DELIVERABLES 2"],
        p_master["TOP 3 PROJECT DELIVERABLES 3"],
    ]

    if "p_name" in kwargs:
        doc.add_paragraph().add_run(kwargs["p_name"]).bold = True
    else:
        doc.add_paragraph().add_run("Top 3 Deliverables").bold = True

    for i, d in enumerate(dels):
        try:  # this is necessary as string data contains NBSP and this removes them
            text_one = d
            text_two = d
            compare_text_new_and_old(text_one, text_two, doc)
        except (TypeError, AttributeError):
            pass

    return doc


def project_report_meta_data(
        doc: Document,
        p_master: Dict,
):
    """Meta data table"""
    # doc.add_section(WD_SECTION_START.NEW_PAGE)
    # paragraph = doc.add_paragraph()
    # paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # paragraph.add_run("Key Info").bold = True

    doc.add_paragraph().add_run("Key Info").bold = True

    if p_master["WLC COMMENTS"] is not None:
        # doc.add_paragraph()
        doc.add_paragraph().add_run("* Total costs comment. " + p_master["WLC COMMENTS"])

    """Costs meta data"""
    t = doc.add_table(rows=1, cols=4)
    hdr_cells = t.rows[0].cells
    hdr_cells[0].text = "ON SCHEDULE:"
    hdr_cells[1].text = str(p_master["PROJECT DEL TO CURRENT TIMINGS ?"])
    hdr_cells[2].text = "ON GMPP:"
    hdr_cells[3].text = str(p_master["GMPP ID: IS THIS PROJECT ON GMPP"])
    row_cells = t.add_row().cells
    row_cells[0].text = "ON BUDGET:"
    row_cells[1].text = str(p_master["PROJECT ON BUDGET?"])
    row_cells[2].text = "TOTAL COST:"
    if p_master["WLC NON GOV"] is None or p_master["WLC NON GOV"] == 0:
        total = p_master["WLC TOTAL"]
    else:
        total = p_master["WLC TOTAL"] + p_master["WLC NON GOV"]
    row_cells[3].text = str(round(total))
    # row_cells = t.add_row().cells
    # row_cells[0].text = "SRO CLEARANCE DATE:"
    # row_cells[1].text = str(p_master["DATE CLEARED BY SRO"])
    # row_cells[2].text = "MOST RECENT PERM SEC REVIEW:"
    # row_cells[3].text = str(p_master["DATE OF MOST RECENT PERM SEC REVIEW"])

    # set column width
    column_widths = (Cm(4), Cm(3), Cm(4), Cm(3))
    set_col_widths(t, column_widths)
    # make column keys bold
    make_columns_bold([t.columns[0], t.columns[2]])
    # change_text_size([t.columns[0], t.columns[1], t.columns[2], t.columns[3]], 10)

    return doc


def dca_narratives(doc: Document, p_master: Dict) -> None:
    doc.add_paragraph()
    # p = doc.add_paragraph()
    # text = "*Red text highlights changes in narratives from last quarter"
    # p.add_run(text).font.color.rgb = RGBColor(255, 0, 0)

    narrative_keys_list = [
        "SRO YOUR ON OFF TRACK ASSESSMENT:",
        "PRIMARY CONCERN & MITIGATING ACTIONS",
        "SHORT UPDATE FOR PM NOTE",
    ]

    headings_list = [
        "SRO assessment narrative",
        "SRO single biggest concern",
        "Short update for PM",
    ]

    for x in range(len(headings_list)):
        try:  # overall try statement relates to data_bridge
            text_one = str(
                p_master[narrative_keys_list[x]]
            )
            try:
                text_two = str(
                    p_master[narrative_keys_list[x]]
                )
            except (KeyError, IndexError):  # index error relates to data_bridge
                text_two = text_one
        except KeyError:
            break

        doc.add_paragraph().add_run(str(headings_list[x])).bold = True
        compare_text_new_and_old(text_one, text_two, doc)


def pm_one_line(doc: Document, p_master: Dict, p_name: str) -> None:
    doc.add_paragraph()
    # p = doc.add_paragraph()
    # text = "*Red text highlights changes in narratives from last quarter"
    # p.add_run(text).font.color.rgb = RGBColor(255, 0, 0)

    narrative_keys_list = [
        "SHORT UPDATE FOR PM NOTE",
    ]

    headings_list = [
        "Short update for PM",
    ]

    for x in range(len(headings_list)):
        try:  # overall try statement relates to data_bridge
            text_one = str(
                p_master[narrative_keys_list[x]]
            )
            try:
                text_two = str(
                    p_master[narrative_keys_list[x]]
                )
            except (KeyError, IndexError):  # index error relates to data_bridge
                text_two = text_one
        except KeyError:
            break

        doc.add_paragraph().add_run(str(p_name)).bold = True
        compare_text_new_and_old(text_one, text_two, doc)

    return doc


def milestone_info_handling(output_list: list, t_list: list) -> list:
    """helper function for handling and cleaning up milestone date generated
    via MilestoneDate class. Removes none type milestone names and non date
    string values"""
    if t_list[1][1] is None or t_list[1][1] == "Project - Business Case End Date":
        pass
    else:
        if isinstance(t_list[2][1], datetime.date):
            return output_list.append(t_list)
        else:
            try:
                d = parser.parse(t_list[2][1], dayfirst=True)
                t_list[3] = ("Date", d.date())
                return output_list.append(t_list)
            # ParserError for non-date string. TypeError for None types
            except (ParserError, TypeError):
                pass


# class MilestoneData:
#     def __init__(
#             self,
#             master: Master,
#             # baseline_type: str = "ipdc_milestones",
#             **kwargs,
#     ):
#         self.master = master
#         self.group = []
#         self.iter_list = []  # iteration list
#         self.kwargs = kwargs
#         # self.baseline_type = baseline_type
#         self.milestone_dict = {}
#         self.sorted_milestone_dict = {}
#         self.max_date = None
#         self.min_date = None
#         self.schedule_change = {}
#         self.schedule_key_last = None
#         self.schedule_key_baseline = None
#         self.get_milestones()
#         self.get_chart_info()
#         # self.calculate_schedule_changes()
#
#     def get_milestones(self) -> None:
#         """
#         Creates project milestone dictionaries for current, last_quarter, and
#         baselines when provided with group and baseline type.
#         """
#         m_dict = {}
#         self.iter_list = get_iter_list(self.kwargs, self.master)
#         for tp in self.iter_list:  # tp time period
#             lower_dict = {}
#             raw_list = []
#             self.group = get_group(self.master, tp, self.kwargs)
#             for project_name in self.group:
#                 project_list = []
#                 p_data = self.master.master_data[0].data[project_name]
#                 # p_data = get_correct_p_data(
#                 #     self.kwargs, self.master, self.baseline_type, project_name, tp
#                 # )
#                 if p_data is None:
#                     continue
#                 # i loops below removes None Milestone names and rejects non-datetime date values.
#                 p = self.master.abbreviations[project_name]["abb"]
#                 for i in range(1, 20):
#                     t = [
#                         ("Project", p),
#                         ("Milestone", p_data["MM" + str(i) + " name"]),
#                         # ("Type", "Approval"),
#                         ("Date", p_data["MM" + str(i) + " date"]),
#                         ("Notes", p_data["MM" + str(i) + " Comment"]),
#                         ("Status", p_data["MM" + str(i) + " status"])
#                     ]
#                     milestone_info_handling(project_list, t)
#
#                 # loop to stop keys names being the same. Done at project level.
#                 # not particularly concise code.
#                 upper_counter_list = []
#                 for entry in project_list:
#                     upper_counter_list.append(entry[1][1])
#                 upper_count = Counter(upper_counter_list)
#                 lower_counter_list = []
#                 for entry in project_list:
#                     if upper_count[entry[1][1]] > 1:
#                         lower_counter_list.append(entry[1][1])
#                         lower_count = Counter(lower_counter_list)
#                         new_milestone_key = (
#                                 entry[1][1] + " (" + str(lower_count[entry[1][1]]) + ")"
#                         )
#                         entry[1] = ("Milestone", new_milestone_key)
#                         raw_list.append(entry)
#                     else:
#                         raw_list.append(entry)
#
#             # puts the list in chronological order
#             sorted_list = sorted(raw_list, key=lambda k: (k[2][1] is None, k[2][1]))
#
#             for r in range(len(sorted_list)):
#                 lower_dict["Milestone " + str(r)] = dict(sorted_list[r])
#
#             m_dict[tp] = lower_dict
#         self.milestone_dict = m_dict
#
#     def get_chart_info(self) -> None:
#         """returns data lists for matplotlib chart"""
#         # Note this code could refactored so that it collects all milestones
#         # reported across current, last and baseline. At the moment it only
#         # uses milestones that are present in the current quarter.
#
#         output_dict = {}
#         for i in self.milestone_dict:
#             key_names = []
#             g_dates = []  # graph dates
#             r_dates = []  # raw dates
#             notes = []
#             status = []
#             for v in self.milestone_dict[self.iter_list[0]].values():
#                 p = None  # project
#                 mn = None  # milestone name
#                 d = None  # date
#                 for x in self.milestone_dict[i].values():
#                     if (
#                             x["Project"] == v["Project"]
#                             and x["Milestone"] == v["Milestone"]
#                     ):
#                         p = x["Project"]
#                         mn = x["Milestone"]
#                         join = p + ", " + mn
#                         # if join not in key_names:  # stop duplicates
#                         key_names.append(join)
#                         d = x["Date"]
#                         g_dates.append(d)
#                         r_dates.append(d)
#                         notes.append(x["Notes"])
#                         status.append(x["Status"])
#                         break
#                 if p is None and mn is None and d is None:
#                     p = v["Project"]
#                     mn = v["Milestone"]
#                     join = p + ", " + mn
#                     # if join not in key_names:
#                     key_names.append(join)
#                     g_dates.append(v["Date"])
#                     r_dates.append(None)
#                     notes.append(None)
#                     status.append(None)
#
#             output_dict[i] = {
#                 "names": key_names,
#                 "g_dates": g_dates,
#                 "r_dates": r_dates,
#                 "notes": notes,
#                 "status": status,
#             }
#
#         self.sorted_milestone_dict = output_dict
#
#     def filter_chart_info(self, **filter_kwargs):
#         # bug handling required in the event that there are no milestones with the filter.
#         # i.e. the filter returns no milestones.
#         filtered_dict = {}
#         if (
#                 "type" in filter_kwargs
#                 and "key" in filter_kwargs
#                 and "dates" in filter_kwargs
#         ):
#             start_date, end_date = zip(*filter_kwargs["dates"])
#             start = parser.parse(start_date, dayfirst=True)
#             end = parser.parse(end_date, dayfirst=True)
#             for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
#                 if v["Type"] in filter_kwargs["type"]:
#                     if v["Milestone"] in filter_kwargs["keys"]:
#                         if start.date() <= filter_kwargs["dates"] <= end.date():
#                             filtered_dict["Milestone " + str(i)] = v
#                             continue
#
#         elif "type" in filter_kwargs and "key" in filter_kwargs:
#             for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
#                 if v["Type"] in filter_kwargs["type"]:
#                     if v["Milestone"] in filter_kwargs["keys"]:
#                         filtered_dict["Milestone " + str(i)] = v
#                         continue
#
#         elif "type" in filter_kwargs and "dates" in filter_kwargs:
#             start_date, end_date = zip(filter_kwargs["dates"])
#             start = parser.parse(start_date[0], dayfirst=True)
#             end = parser.parse(end_date[0], dayfirst=True)
#             for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
#                 if v["Type"] in filter_kwargs["type"]:
#                     if start.date() <= v["Date"] <= end.date():
#                         filtered_dict["Milestone " + str(i)] = v
#                         continue
#
#         elif "key" in filter_kwargs and "dates" in filter_kwargs:
#             start_date, end_date = zip(filter_kwargs["dates"])
#             start = parser.parse(start_date[0], dayfirst=True)
#             end = parser.parse(end_date[0], dayfirst=True)
#             for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
#                 if v["Milestone"] in filter_kwargs["key"]:
#                     if start.date() <= v["Date"] <= end.date():
#                         filtered_dict["Milestone " + str(i)] = v
#                         continue
#
#         elif "type" in filter_kwargs:
#             for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
#                 if v["Type"] in filter_kwargs["type"]:
#                     filtered_dict["Milestone " + str(i)] = v
#                     continue
#
#         elif "key" in filter_kwargs:
#             for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
#                 if v["Milestone"] in filter_kwargs["key"]:
#                     filtered_dict["Milestone " + str(i)] = v
#                     continue
#
#         elif "dates" in filter_kwargs:
#             start_date, end_date = zip(filter_kwargs["dates"])
#             start = parser.parse(start_date[0], dayfirst=True)
#             end = parser.parse(end_date[0], dayfirst=True)
#             for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
#                 if start.date() <= v["Date"] <= end.date():
#                     filtered_dict["Milestone " + str(i)] = v
#                     continue
#
#         output_dict = {}
#         for dict in self.milestone_dict.keys():
#             if dict == self.iter_list[0]:
#                 output_dict[dict] = filtered_dict
#             else:
#                 output_dict[dict] = self.milestone_dict[dict]
#
#         self.milestone_dict = output_dict
#         self.get_chart_info()
#
#     def calculate_schedule_changes(self) -> None:
#         """calculates the changes in project schedules. If standard key for calculation
#         not available it using the best next one available"""
#
#         self.filter_chart_info(milestone_type=["Delivery", "Approval"])
#         m_dict_keys = list(self.milestone_dict.keys())
#
#         def schedule_info(
#                 project_name: str,
#                 other_key_list: List[str],
#                 c_key_list: List[str],
#                 miles_dict: dict,
#                 dict_l_current: str,
#                 dict_l_other: str,
#         ):
#             output_dict = {}
#             schedule_info = []
#             for key in reversed(other_key_list):
#                 if key in c_key_list:
#                     sop = get_milestone_date(
#                         project_name, miles_dict, dict_l_other, " Start of Project"
#                     )
#                     if sop is None:
#                         sop = get_milestone_date(
#                             project_name, miles_dict, dict_l_current, other_key_list[0]
#                         )
#                         schedule_info.append(("start key", other_key_list[0]))
#                     else:
#                         schedule_info.append(("start key", " Start of Project"))
#                     schedule_info.append(("start", sop))
#                     schedule_info.append(("end key", key))
#                     date = get_milestone_date(
#                         project_name, miles_dict, dict_l_current, key
#                     )
#                     schedule_info.append(("end current date", date))
#                     other_date = get_milestone_date(
#                         project_name, miles_dict, dict_l_other, key
#                     )
#                     schedule_info.append(("end other date", other_date))
#                     project_length = (other_date - sop).days
#                     schedule_info.append(("project length", project_length))
#                     change = (date - other_date).days
#                     schedule_info.append(("change", change))
#                     p_change = int((change / project_length) * 100)
#                     schedule_info.append(("percent change", p_change))
#                     output_dict[dict_l_other] = dict(schedule_info)
#                     break
#
#             return output_dict
#
#         output_dict = {}
#         for project_name in self.group:
#             project_name = self.master.abbreviations[project_name]
#             current_key_list = []
#             last_key_list = []
#             baseline_key_list = []
#             for key in self.key_names:
#                 try:
#                     p = key.split(",")[0]
#                     milestone_key = key.split(",")[1]
#                     if project_name == p:
#                         if milestone_key != " Project - Business Case End Date":
#                             current_key_list.append(milestone_key)
#                 except IndexError:
#                     # patch of single project group. In this instance the project name
#                     # is removed from the key_name via remove_project_name function as
#                     # part of get chart info.
#                     if len(self.group) == 1:
#                         current_key_list.append(" " + key)
#             for last_key in self.key_names_last:
#                 p = last_key.split(",")[0]
#                 milestone_key_last = last_key.split(",")[1]
#                 if project_name == p:
#                     if milestone_key_last != " Project - Business Case End Date":
#                         last_key_list.append(milestone_key_last)
#             for baseline_key in self.key_names_baseline:
#                 p = baseline_key.split(",")[0]
#                 milestone_key_baseline = baseline_key.split(",")[1]
#                 if project_name == p:
#                     if (
#                             milestone_key_baseline
#                             != " Project - Business Case End Date"
#                             # and milestone_key_baseline != " Project End Date"
#                     ):
#                         baseline_key_list.append(milestone_key_baseline)
#
#             b_dict = schedule_info(
#                 project_name,
#                 baseline_key_list,
#                 current_key_list,
#                 self.milestone_dict,
#                 m_dict_keys[0],
#                 m_dict_keys[2],
#             )
#             l_dict = schedule_info(
#                 project_name,
#                 last_key_list,
#                 current_key_list,
#                 self.milestone_dict,
#                 m_dict_keys[0],
#                 m_dict_keys[1],
#             )
#             lower_dict = {**b_dict, **l_dict}
#
#             output_dict[project_name] = lower_dict
#
#         self.schedule_change = output_dict


class CentralSupportData:
    def __init__(
            self,
            master: Master,
            # baseline_type: str = "ipdc_milestones",
            **kwargs,
    ):
        self.master = master
        self.group = []
        self.iter_list = []  # iteration list
        self.kwargs = kwargs
        # self.baseline_type = baseline_type
        self.milestone_dict = {}
        self.sorted_milestone_dict = {}
        self.max_date = None
        self.min_date = None
        self.schedule_change = {}
        self.schedule_key_last = None
        self.schedule_key_baseline = None
        self.get_milestones()
        self.get_chart_info()
        # self.calculate_schedule_changes()

    def get_milestones(self) -> None:
        """
        Creates project milestone dictionaries for current, last_quarter, and
        baselines when provided with group and baseline type.
        """
        sp_dict = {}
        self.iter_list = get_iter_list(self.kwargs, self.master)
        for tp in self.iter_list:  # tp time period
            lower_dict = {}
            raw_list = []
            self.group = get_group(self.master, tp, self.kwargs)
            for project_name in self.group:
                project_list = []
                p_data = self.master.master_data[0].data[project_name]
                # p_data = get_correct_p_data(
                #     self.kwargs, self.master, self.baseline_type, project_name, tp
                # )
                if p_data is None:
                    continue
                # i loops below removes None Milestone names and rejects non-datetime date values.
                p = self.master.abbreviations[project_name]["abb"]
                for i in range(1, 20):
                    t = [
                        ("Project", p),
                        ("Requirement", p_data["R" + str(i) + " name"]),
                        # ("Type", "Approval"),
                        ("Date", p_data["R" + str(i) + " needed by"]),
                        ("Escalated", p_data["R" + str(i) + " escalated to"]),
                        ("Type", p_data["R" + str(i) + " type"]),
                        ("Secured", p_data["R" + str(i) + " secured"]),
                    ]
                    milestone_info_handling(project_list, t)

                # loop to stop keys names being the same. Done at project level.
                # not particularly concise code.
                upper_counter_list = []
                for entry in project_list:
                    upper_counter_list.append(entry[1][1])
                upper_count = Counter(upper_counter_list)
                lower_counter_list = []
                for entry in project_list:
                    if upper_count[entry[1][1]] > 1:
                        lower_counter_list.append(entry[1][1])
                        lower_count = Counter(lower_counter_list)
                        new_require_key = (
                                entry[1][1] + " (" + str(lower_count[entry[1][1]]) + ")"
                        )
                        entry[1] = ("Milestone", new_require_key)
                        raw_list.append(entry)
                    else:
                        raw_list.append(entry)

            # puts the list in chronological order
            sorted_list = sorted(raw_list, key=lambda k: (k[3][1] is None, k[3][1]))

            for r in range(len(sorted_list)):
                lower_dict["Requirement " + str(r)] = dict(sorted_list[r])

            sp_dict[tp] = lower_dict
        self.milestone_dict = sp_dict

    def get_chart_info(self) -> None:
        """returns data lists for matplotlib chart"""
        # Note this code could refactored so that it collects all milestones
        # reported across current, last and baseline. At the moment it only
        # uses milestones that are present in the current quarter.

        output_dict = {}
        for i in self.milestone_dict:
            key_names = []
            g_dates = []  # graph dates
            r_dates = []  # raw dates
            escalated = []
            type = []
            secured = []
            for v in self.milestone_dict[self.iter_list[0]].values():
                p = None  # project
                mn = None  # milestone name
                d = None  # date
                for x in self.milestone_dict[i].values():
                    if (
                            x["Project"] == v["Project"]
                            and x["Requirement"] == v["Requirement"]
                    ):
                        p = x["Project"]
                        mn = x["Requirement"]
                        join = p + ", " + mn
                        # if join not in key_names:  # stop duplicates
                        key_names.append(join)
                        d = x["Date"]
                        g_dates.append(d)
                        r_dates.append(d)
                        escalated.append(x["Escalated"])
                        type.append(x["Type"])
                        secured.append(x["Secured"])
                        break
                if p is None and mn is None and d is None:
                    p = v["Project"]
                    mn = v["Requirement"]
                    join = p + ", " + mn
                    # if join not in key_names:
                    key_names.append(join)
                    g_dates.append(v["Date"])
                    r_dates.append(None)
                    escalated.append(None)
                    type.append(None)
                    secured.append(None)

            output_dict[i] = {
                "names": key_names,
                "g_dates": g_dates,
                "r_dates": r_dates,
                "escalated": escalated,
                "type": type,
                "secured": secured,
            }

        self.sorted_milestone_dict = output_dict

    def filter_chart_info(self, **filter_kwargs):
        # bug handling required in the event that there are no milestones with the filter.
        # i.e. the filter returns no milestones.
        filtered_dict = {}
        if (
                "type" in filter_kwargs
                and "key" in filter_kwargs
                and "dates" in filter_kwargs
        ):
            start_date, end_date = zip(*filter_kwargs["dates"])
            start = parser.parse(start_date, dayfirst=True)
            end = parser.parse(end_date, dayfirst=True)
            for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
                if v["Type"] in filter_kwargs["type"]:
                    if v["Milestone"] in filter_kwargs["keys"]:
                        if start.date() <= filter_kwargs["dates"] <= end.date():
                            filtered_dict["Milestone " + str(i)] = v
                            continue

        elif "type" in filter_kwargs and "key" in filter_kwargs:
            for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
                if v["Type"] in filter_kwargs["type"]:
                    if v["Milestone"] in filter_kwargs["keys"]:
                        filtered_dict["Milestone " + str(i)] = v
                        continue

        elif "type" in filter_kwargs and "dates" in filter_kwargs:
            start_date, end_date = zip(filter_kwargs["dates"])
            start = parser.parse(start_date[0], dayfirst=True)
            end = parser.parse(end_date[0], dayfirst=True)
            for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
                if v["Type"] in filter_kwargs["type"]:
                    if start.date() <= v["Date"] <= end.date():
                        filtered_dict["Milestone " + str(i)] = v
                        continue

        elif "key" in filter_kwargs and "dates" in filter_kwargs:
            start_date, end_date = zip(filter_kwargs["dates"])
            start = parser.parse(start_date[0], dayfirst=True)
            end = parser.parse(end_date[0], dayfirst=True)
            for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
                if v["Milestone"] in filter_kwargs["key"]:
                    if start.date() <= v["Date"] <= end.date():
                        filtered_dict["Milestone " + str(i)] = v
                        continue

        elif "type" in filter_kwargs:
            for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
                if v["Type"] in filter_kwargs["type"]:
                    filtered_dict["Milestone " + str(i)] = v
                    continue

        elif "key" in filter_kwargs:
            for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
                if v["Milestone"] in filter_kwargs["key"]:
                    filtered_dict["Milestone " + str(i)] = v
                    continue

        elif "dates" in filter_kwargs:
            start_date, end_date = zip(filter_kwargs["dates"])
            start = parser.parse(start_date[0], dayfirst=True)
            end = parser.parse(end_date[0], dayfirst=True)
            for i, v in enumerate(self.milestone_dict[self.iter_list[0]].values()):
                if start.date() <= v["Date"] <= end.date():
                    filtered_dict["Milestone " + str(i)] = v
                    continue

        output_dict = {}
        for dict in self.milestone_dict.keys():
            if dict == self.iter_list[0]:
                output_dict[dict] = filtered_dict
            else:
                output_dict[dict] = self.milestone_dict[dict]

        self.milestone_dict = output_dict
        self.get_chart_info()

    def calculate_schedule_changes(self) -> None:
        """calculates the changes in project schedules. If standard key for calculation
        not available it using the best next one available"""

        self.filter_chart_info(milestone_type=["Delivery", "Approval"])
        m_dict_keys = list(self.milestone_dict.keys())

        def schedule_info(
                project_name: str,
                other_key_list: List[str],
                c_key_list: List[str],
                miles_dict: dict,
                dict_l_current: str,
                dict_l_other: str,
        ):
            output_dict = {}
            schedule_info = []
            for key in reversed(other_key_list):
                if key in c_key_list:
                    sop = get_milestone_date(
                        project_name, miles_dict, dict_l_other, " Start of Project"
                    )
                    if sop is None:
                        sop = get_milestone_date(
                            project_name, miles_dict, dict_l_current, other_key_list[0]
                        )
                        schedule_info.append(("start key", other_key_list[0]))
                    else:
                        schedule_info.append(("start key", " Start of Project"))
                    schedule_info.append(("start", sop))
                    schedule_info.append(("end key", key))
                    date = get_milestone_date(
                        project_name, miles_dict, dict_l_current, key
                    )
                    schedule_info.append(("end current date", date))
                    other_date = get_milestone_date(
                        project_name, miles_dict, dict_l_other, key
                    )
                    schedule_info.append(("end other date", other_date))
                    project_length = (other_date - sop).days
                    schedule_info.append(("project length", project_length))
                    change = (date - other_date).days
                    schedule_info.append(("change", change))
                    p_change = int((change / project_length) * 100)
                    schedule_info.append(("percent change", p_change))
                    output_dict[dict_l_other] = dict(schedule_info)
                    break

            return output_dict

        output_dict = {}
        for project_name in self.group:
            project_name = self.master.abbreviations[project_name]
            current_key_list = []
            last_key_list = []
            baseline_key_list = []
            for key in self.key_names:
                try:
                    p = key.split(",")[0]
                    milestone_key = key.split(",")[1]
                    if project_name == p:
                        if milestone_key != " Project - Business Case End Date":
                            current_key_list.append(milestone_key)
                except IndexError:
                    # patch of single project group. In this instance the project name
                    # is removed from the key_name via remove_project_name function as
                    # part of get chart info.
                    if len(self.group) == 1:
                        current_key_list.append(" " + key)
            for last_key in self.key_names_last:
                p = last_key.split(",")[0]
                milestone_key_last = last_key.split(",")[1]
                if project_name == p:
                    if milestone_key_last != " Project - Business Case End Date":
                        last_key_list.append(milestone_key_last)
            for baseline_key in self.key_names_baseline:
                p = baseline_key.split(",")[0]
                milestone_key_baseline = baseline_key.split(",")[1]
                if project_name == p:
                    if (
                            milestone_key_baseline
                            != " Project - Business Case End Date"
                            # and milestone_key_baseline != " Project End Date"
                    ):
                        baseline_key_list.append(milestone_key_baseline)

            b_dict = schedule_info(
                project_name,
                baseline_key_list,
                current_key_list,
                self.milestone_dict,
                m_dict_keys[0],
                m_dict_keys[2],
            )
            l_dict = schedule_info(
                project_name,
                last_key_list,
                current_key_list,
                self.milestone_dict,
                m_dict_keys[0],
                m_dict_keys[1],
            )
            lower_dict = {**b_dict, **l_dict}

            output_dict[project_name] = lower_dict

        self.schedule_change = output_dict


def print_out_project_milestones(
        doc: Document,
        milestones: MilestoneData,
) -> Document:
    # doc.add_section(WD_SECTION_START.NEW_PAGE)
    # table heading
    # ab = milestones.master.abbreviations[project_name]["abb"]
    doc.add_paragraph().add_run("Milestones").bold = True

    table = doc.add_table(rows=1, cols=4)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Milestone"
    hdr_cells[1].text = "Date"
    hdr_cells[2].text = "Status"
    # hdr_cells[3].text = "Change from baseline"
    hdr_cells[3].text = "Notes"

    if not milestones.sorted_milestone_dict[milestones.iter_list[0]]["names"]:
        doc.add_paragraph().add_run("No milestones reported")
    else:
        for i, m in enumerate(
                milestones.sorted_milestone_dict[milestones.iter_list[0]]["names"]
        ):
            row_cells = table.add_row().cells
            if len(milestones.group) == 1:
                no_name = m.split(",")[1]
                row_cells[0].text = no_name
            else:
                row_cells[0].text = m
            row_cells[1].text = milestones.sorted_milestone_dict[milestones.iter_list[0]][
                "r_dates"
            ][i].strftime("%d/%m/%Y")
            row_cells[2].text = milestones.sorted_milestone_dict[milestones.iter_list[0]][
                "status"][i]
            # try:
            #     row_cells[2].text = plus_minus_days(
            #         (
            #                 milestones.sorted_milestone_dict[milestones.iter_list[0]][
            #                     "r_dates"
            #                 ][i]
            #                 - milestones.sorted_milestone_dict[milestones.iter_list[1]][
            #                     "r_dates"
            #                 ][i]
            #         ).days
            #     )
            # except TypeError:
            #     row_cells[2].text = "Not reported"
            # try:
            #     row_cells[3].text = plus_minus_days(
            #         (
            #                 milestones.sorted_milestone_dict[milestones.iter_list[0]][
            #                     "r_dates"
            #                 ][i]
            #                 - milestones.sorted_milestone_dict[milestones.iter_list[2]][
            #                     "r_dates"
            #                 ][i]
            #         ).days
            #     )
            # except TypeError:
            #     row_cells[3].text = "Not reported"
            try:
                row_cells[3].text = milestones.sorted_milestone_dict[
                    milestones.iter_list[0]
                ]["notes"][i]
                paragraph = row_cells[3].paragraphs[0]
                run = paragraph.runs
                font = run[0].font
                font.size = Pt(8)  # font size = 8
            except TypeError:
                pass

        table.style = "Table Grid"

        # column widths
        column_widths = (Cm(6), Cm(2.6), Cm(2.6), Cm(10))
        set_col_widths(table, column_widths)
        # make_columns_bold([table.columns[0], table.columns[3]])  # make keys bold
        # make_text_red([table.columns[1], table.columns[4]])  # make 'not reported red'

        # make_rows_bold(
        #     [table.rows[0]]
        # )  # makes top of table bold. Found function on stack overflow.
    return doc


def print_out_central_support(
        doc: Document,
        centrals: CentralSupportData,
) -> Document:
    # doc.add_section(WD_SECTION_START.NEW_PAGE)
    # table heading
    # ab = milestones.master.abbreviations[project_name]["abb"]
    doc.add_paragraph()
    doc.add_paragraph().add_run("Central govnt support requirements").bold = True

    if not centrals.sorted_milestone_dict[centrals.iter_list[0]]["names"]:
        doc.add_paragraph().add_run("No requirements reported")
    else:
        table = doc.add_table(rows=1, cols=5)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Requirement"
        hdr_cells[1].text = "Need by"
        hdr_cells[2].text = "Escalated to"
        hdr_cells[3].text = "Type"
        hdr_cells[4].text = "Secured"

        for i, m in enumerate(
                centrals.sorted_milestone_dict[centrals.iter_list[0]]["names"]
        ):
            row_cells = table.add_row().cells
            if len(centrals.group) == 1:
                no_name = m.split(",")[1]
                row_cells[0].text = no_name
            else:
                row_cells[0].text = m
            row_cells[1].text = centrals.sorted_milestone_dict[centrals.iter_list[0]][
                "r_dates"
            ][i].strftime("%d/%m/%Y")
            row_cells[2].text = str(centrals.sorted_milestone_dict[centrals.iter_list[0]][
                "escalated"][i])
            row_cells[3].text = str(centrals.sorted_milestone_dict[centrals.iter_list[0]][
                "type"][i])
            row_cells[4].text = str(centrals.sorted_milestone_dict[centrals.iter_list[0]][
                "secured"][i])
            # try:
            #     row_cells[2].text = plus_minus_days(
            #         (
            #                 milestones.sorted_milestone_dict[milestones.iter_list[0]][
            #                     "r_dates"
            #                 ][i]
            #                 - milestones.sorted_milestone_dict[milestones.iter_list[1]][
            #                     "r_dates"
            #                 ][i]
            #         ).days
            #     )
            # except TypeError:
            #     row_cells[2].text = "Not reported"
            # try:
            #     row_cells[3].text = plus_minus_days(
            #         (
            #                 milestones.sorted_milestone_dict[milestones.iter_list[0]][
            #                     "r_dates"
            #                 ][i]
            #                 - milestones.sorted_milestone_dict[milestones.iter_list[2]][
            #                     "r_dates"
            #                 ][i]
            #         ).days
            #     )
            # except TypeError:
            #     row_cells[3].text = "Not reported"
            # try:
            #     row_cells[3].text = milestones.sorted_milestone_dict[
            #         milestones.iter_list[0]
            #     ]["notes"][i]
            #     paragraph = row_cells[3].paragraphs[0]
            #     run = paragraph.runs
            #     font = run[0].font
            #     font.size = Pt(8)  # font size = 8
            # except TypeError:
            #     pass

        table.style = "Table Grid"

        # column widths
        column_widths = (Cm(6), Cm(2.5), Cm(2.5), Cm(2.5), Cm(2.5))
        set_col_widths(table, column_widths)
    # make_columns_bold([table.columns[0], table.columns[3]])  # make keys bold
    # make_text_red([table.columns[1], table.columns[4]])  # make 'not reported red'

    # make_rows_bold(
    #     [table.rows[0]]
    # )  # makes top of table bold. Found function on stack overflow.
    return doc


class DandelionData:
    def __init__(self, master: Master, **kwargs):
        self.master = master
        self.kwargs = kwargs
        self.baseline_type = "ipdc_costs"
        self.group = []
        self.iter_list = []
        self.d_data = {}
        self.get_data()

    def get_data(self):
        self.iter_list = get_iter_list(self.kwargs, self.master)
        for tp in self.iter_list:  # although tp is iterated only one can be handled for now.
            #  for dandelion need groups of groups.
            if "group" in self.kwargs:
                self.group = self.kwargs["group"]
            elif "stage" in self.kwargs:
                self.group = self.kwargs["stage"]

            if len(self.group) == 5:
                g_ang_l = [260, 310, 360, 50, 100]  # group angle list
            if len(self.group) == 4:
                g_ang_l = [260, 326, 32, 100]
            if len(self.group) == 3:
                g_ang_l = [280, 360, 80]
            if len(self.group) == 2:
                g_ang_l = [290, 70]
            if len(self.group) == 1:
                pass
            g_d = {}  # group dictionary. first outer circle.
            l_g_d = {}  # lower group dictionary

            pf_wlc = get_dandelion_type_total(
                self.master, tp, self.group, self.kwargs
            )  # portfolio wlc
            if "pc" in self.kwargs:  # pc portfolio colour
                pf_colour = COLOUR_DICT[self.kwargs["pc"]]
                pf_colour_edge = COLOUR_DICT[self.kwargs["pc"]]
            else:
                pf_colour = "#FFFFFF"
                pf_colour_edge = "grey"
            pf_text = "Portfolio\n" + dandelion_number_text(
                pf_wlc
            )  # option to specify pf name

            ## center circle
            g_d["portfolio"] = {
                "axis": (0, 0),
                "r": math.sqrt(pf_wlc),
                "colour": pf_colour,
                "text": pf_text,
                "fill": "solid",
                "ec": pf_colour_edge,
                "alignment": ("center", "center"),
            }

            ## first outer circle
            for i, g in enumerate(self.group):
                self.kwargs["group"] = [g]
                g_wlc = get_dandelion_type_total(self.master, tp, g, self.kwargs)
                if len(self.group) > 1:
                    y_axis = 0 + (
                            (math.sqrt(pf_wlc) * 3.25) * math.sin(math.radians(g_ang_l[i]))
                    )
                    x_axis = 0 + (math.sqrt(pf_wlc) * 2.75) * math.cos(
                        math.radians(g_ang_l[i])
                    )
                    g_text = g + "\n" + dandelion_number_text(g_wlc)  # group text
                    if g_wlc == 0:
                        g_wlc = pf_wlc/20
                    g_d[g] = {
                        "axis": (y_axis, x_axis),
                        "r": math.sqrt(g_wlc),
                        "wlc": g_wlc,
                        "colour": "#FFFFFF",
                        "text": g_text,
                        "fill": "dashed",
                        "ec": "grey",
                        "alignment": ("center", "center"),
                        "angle": g_ang_l[i],
                    }

                else:
                    g_d = {}
                    pf_wlc = g_wlc * 3
                    g_text = g + "\n" + dandelion_number_text(g_wlc)  # group text
                    if g_wlc == 0:
                        g_wlc = 5
                    g_d[g] = {
                        "axis": (0, 0),
                        "r": math.sqrt(g_wlc),
                        "wlc": g_wlc,
                        "colour": "#FFFFFF",
                        "text": g_text,
                        "fill": "dashed",
                        "ec": "grey",
                        "alignment": ("center", "center"),
                    }

            ## second outer circle
            for i, g in enumerate(self.group):
                self.kwargs["group"] = [g]
                group = get_group(self.master, tp, self.kwargs)  # lower group
                p_list = []
                for p in group:
                    self.kwargs["group"] = [p]
                    p_value = get_dandelion_type_total(
                        self.master, tp, p, self.kwargs
                    )  # project wlc
                    p_list.append((p_value, p))
                l_g_d[g] = list(reversed(sorted(p_list)))

            for g in self.group:
                g_wlc = g_d[g]["wlc"]
                g_radius = g_d[g]["r"]
                g_y_axis = g_d[g]["axis"][0]  # group y axis
                g_x_axis = g_d[g]["axis"][1]  # group x axis
                try:
                    p_values_list, p_list = zip(*l_g_d[g])
                except ValueError:  # handles no projects in l_g_d list
                    continue
                if len(p_list) > 3 or len(self.group) == 1:
                    ang_l = cal_group_angle(360, p_list, all=True)
                else:
                    if len(p_list) == 1:
                        ang_l = [g_d[g]["angle"]]
                    if len(p_list) == 2:
                        ang_l = [g_d[g]["angle"], g_d[g]["angle"] + 60]
                    if len(p_list) == 3:
                        ang_l = [g_d[g]["angle"], g_d[g]["angle"] + 60, g_d[g]["angle"] + 120]

                for i, p in enumerate(p_list):
                    p_value = p_values_list[i]
                    p_data = get_correct_p_data(
                        self.kwargs, self.master, self.baseline_type, p, tp
                    )
                    # change confidence type here
                    # SRO Schedule Confidence
                    # Departmental DCA
                    # SRO Benefits RAG
                    # rag = p_data["Departmental DCA"]
                    colour = COLOUR_DICT[convert_rag_text(None)]  # no rags for 250
                    project_text = (
                            self.master.abbreviations[p]["abb"]
                            + "\n"
                            + dandelion_number_text(p_value)
                    )
                    if p_value == 0:
                        p_value = 200
                    if p in self.master.dft_groups[tp]["GMPP"]:
                        edge_colour = "#000000"  # edge of bubble
                    else:
                        edge_colour = colour

                    # multi = math.sqrt(pf_wlc/g_wlc)  # multiplier
                    # multi = (1 - (g_wlc / pf_wlc)) * 3
                    try:
                        if len(p_list) >= 14:
                            multi = ((pf_wlc / g_wlc) ** (1.0 / 2.0))  # square root
                        else:
                            multi = (pf_wlc / g_wlc) ** (1.0 / 3.0)  # cube root
                        p_y_axis = g_y_axis + (g_radius * multi) * math.sin(
                            math.radians(ang_l[i])
                        )
                        p_x_axis = g_x_axis + (g_radius * multi) * math.cos(
                            math.radians(ang_l[i])
                        )
                    except ZeroDivisionError:
                        p_y_axis = g_y_axis + 100 * math.sin(math.radians(ang_l[i]))
                        p_x_axis = g_x_axis + 100 * math.cos(math.radians(ang_l[i]))

                    if 185 >= ang_l[i] >= 175:
                        text_angle = ("center", "top")
                    if 5 >= ang_l[i] or 355 <= ang_l[i]:
                        text_angle = ("center", "bottom")
                    if 174 >= ang_l[i] >= 6:
                        text_angle = ("left", "center")
                    if 354 >= ang_l[i] >= 186:
                        text_angle = ("right", "center")

                    try:
                        t_multi = (g_wlc / p_value) ** (1.0 / 4.0)
                        # t_multi = (1 - (p_value/g_wlc)) * 2  # text multiplier
                    except ZeroDivisionError:
                        t_multi = 1
                    yx_text_position = (
                        p_y_axis
                        + (math.sqrt(p_value) * t_multi)
                        * math.sin(math.radians(ang_l[i])),
                        p_x_axis
                        + (math.sqrt(p_value) * t_multi)
                        * math.cos(math.radians(ang_l[i])),
                    )

                    g_d[p] = {
                        "axis": (p_y_axis, p_x_axis),
                        "r": math.sqrt(p_value),
                        "wlc": p_value,
                        "colour": colour,
                        "text": project_text,
                        "fill": "solid",
                        "ec": "grey",
                        "alignment": text_angle,
                        "tp": yx_text_position,
                    }

        self.d_data = g_d


def get_dandelion_type_total(
        master: Master, tp: str, g: str or List[str], kwargs
) -> int or str:  # Note no **kwargs as existing kwargs dict passed in
    if "type" in kwargs:
        if kwargs["type"] == "remaining":
            cost = CostData(master, quarter=[tp], group=[g])  # group costs data
            return cost.c_totals[tp]["prof"] + cost.c_totals[tp]["unprof"]
        if kwargs["type"] == "spent":
            cost = CostData(master, quarter=[tp], group=[g])  # group costs data
            return cost.c_totals[tp]["spent"]
        # if kwargs["type"] == "benefits":
        #     benefits = BenefitsData(master, quarter=[tp], group=[g])
        #     return benefits.b_totals[tp]["total"]

    else:
        cost = CostData(master, **kwargs)  # group costs data
        return cost.c_totals[tp]["total"]

