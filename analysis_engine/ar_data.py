"""place for all code used to make annual report summaries"""
from datetime import date
from typing import Dict, Union

from datamaps.api import project_data_from_master
from docx.enum.section import WD_SECTION_START

from analysis_engine.data import (
    # root_path,
    open_word_doc,
    # get_project_information,
    wd_heading,
    compare_text_new_and_old,
)

from analysis_engine.core_data import get_project_information

from docx import Document


def get_gmpp_ar_data(
    confi_path: Path,
) -> Dict:
    try:
        config = configparser.ConfigParser()
        config.read(confi_path)
        path_dict = {
            "gmpp_ar_master": config["GMPP ANNUAL REPORT"]["ar_master_data"],
        }
    except:
        logger.critical(
            "Configuration file issue. Please check and make sure it's correct."
        )
        sys.exit(1)

    return path_dict


def get_ar_data(file_name):
    return project_data_from_master(
        root_path / "core_data/{}.xlsx".format(file_name), 4, 2021
    )


def ar_run_p_reports(
    data: dict,
    # project_list: list,
) -> None:
    report_doc = open_word_doc(root_path / "input/summary_temp.docx")

    for i, p in enumerate(data.projects):
        if i != 0:
            report_doc.add_section(WD_SECTION_START.NEW_PAGE)  # new page
        print("Compiling " + p)
        # qrt = make_file_friendly(str(master.ma.quarter))
        output = ar_compile_p_report(
            data,
            report_doc,
            # project_list,
            p,
            i,
        )
        # abb = project_info[p]["Abbreviations"]
    output.save(root_path / "output/annual_report_summaries.docx")


def ar_compile_p_report(
    data: dict,
    doc: Document,
    # project_info: Dict[str, Union[str, int, date, float]],
    project_name: str,
    no: int,
) -> Document:
    # print(project_name)
    # if no != 0:
    #     wd_heading(doc, group=[project_name], delete=True)
    wd_heading(doc, group=[project_name])
    # wd_heading(doc, project_info, project_name, data_type="ar")
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
    "Date of the latest approved HMT Treasury Approval point (sent to PAC)",
    "Whole Life Costs (£m) latest-approved HMT TAP (Information sent to PAC)",
    "Departmental Narrative on WLC variance between the department baseline and the HMT latest-approved Baseline (Information sent to PAC)",
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

        v = data[project_name][headings_list[x]]

        if isinstance(v, date):
            text_one = v.strftime("%d/%m/%Y")
        else:
            text_one = str(v)
            # try:
            #     text_two = str(data[project_name][headings_list[x]])
            # except (KeyError, IndexError):  # index error relates to data_bridge
            text_two = text_one

        if (
            headings_list[x]
            == "Date of the latest approved HMT Treasury Approval point (sent to PAC)"
        ):
            text_one = (
                str(data[project_name]["Type of the latest-approved HMT TAP"])
                + ": "
                + text_one
            )
            text_two = text_one

        doc.add_paragraph().add_run(str(headings_list[x])).bold = True

        compare_text_new_and_old(text_one, text_two, doc)


# data = get_ar_data("core_data/2122_ar_master.xlsx")
# ar_run_p_reports(data)
