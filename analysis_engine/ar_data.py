"""place for all code used to make annual report summaries"""
from datetime import date
from typing import Dict, Union

from datamaps.api import project_data_from_master
from docx.enum.section import WD_SECTION_START

from analysis_engine.data import (
    root_path,
    open_word_doc,
    get_project_information,
    wd_heading,
    compare_text_new_and_old,
)

from docx import Document


def get_ar_data():
    return project_data_from_master(root_path / "core_data/ar_master.xlsx", 4, 2020)


def ar_run_p_reports(data: dict, project_info: dict) -> None:
    report_doc = open_word_doc(root_path / "input/summary_temp.docx")

    for i, p in enumerate(data.projects):
        if i != 0:
            report_doc.add_section(WD_SECTION_START.NEW_PAGE)  # new page
        # print("Compiling summary for " + p)
        # qrt = make_file_friendly(str(master.ma.quarter))
        output = ar_compile_p_report(data, report_doc, get_project_information(), p)
        # abb = project_info[p]["Abbreviations"]
    output.save(root_path / "output/annual_report_summaries.docx")


def ar_compile_p_report(
    data: dict,
    doc: Document,
    project_info: Dict[str, Union[str, int, date, float]],
    project_name: str,
) -> Document:
    wd_heading(doc, project_info, project_name, data_type="ar")
    ar_narratives(doc, data, project_name, AR_NARRATIVES)

    return doc


AR_NARRATIVES = [
    "Project Description",
    "IPA Delivery Confidence Assessment (DCA)",
    "Departmental DCA Narrative",
    "Latest Approved Project End Date",
    "Departmental Schedule Narrative",
    # "Financial Year Variance",
    "Departmental Financial Year Narrative",
    "Whole Life Cost (WLC)",
    "Departmental WLC Narrative",
]


def ar_narratives(
    doc: Document,
    data: dict,
    project_name: str,
    headings_list: list,
) -> None:
    """Places all narratives into document and checks for differences between
    current and last quarter"""

    for x in range(len(headings_list)):
        try:  # overall try statement relates to data_bridge
            text_one = str(data[project_name][headings_list[x]])
            try:
                text_two = str(data[project_name][headings_list[x]])
            except (KeyError, IndexError):  # index error relates to data_bridge
                text_two = text_one
        except KeyError:
            break

        doc.add_paragraph().add_run(str(headings_list[x])).bold = True

        compare_text_new_and_old(text_one, text_two, doc)
