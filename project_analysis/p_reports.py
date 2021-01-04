"""
New code for compiling individual project reports.
"""
import os
from datetime import date
from typing import Dict, Union

from docx import Document
from data_mgmt.data import (
    root_path,
    CostData,
    Projects,
    get_master_data,
    Master,
    open_word_doc,
    wd_heading,
    key_contacts,
    get_project_information,
    dca_table,
    dca_narratives,
    put_matplotlib_fig_into_word,
    cost_profile_graph,
    change_word_doc_landscape,
    BenefitsData,
    total_costs_benefits_bar_chart,
    FIGURE_STYLE, MilestoneData, milestone_chart, project_report_meta_data,
)


def compile_report(
    doc: Document,
    project_info: Dict[str, Union[str, int, date, float]],
    master: Master,
    project_name: str,
) -> Document:
    wd_heading(doc, project_info, project_name)
    key_contacts(doc, master, project_name)
    dca_table(doc, master, project_name)
    dca_narratives(doc, master, project_name)
    costs = CostData(master, project_name)
    benefits = BenefitsData(master, project_name)
    milestones = MilestoneData(master, project_name)
    project_report_meta_data(doc, costs, milestones, benefits, project_name)
    change_word_doc_landscape(doc)
    # fig_style = FIGURE_STYLE[2]
    total_profile = total_costs_benefits_bar_chart(costs, benefits)
    put_matplotlib_fig_into_word(doc, total_profile)
    cost_profile = cost_profile_graph(costs)
    put_matplotlib_fig_into_word(doc, cost_profile)
    milestones.filter_chart_info(start_date="1/1/2021", end_date="20/12/2022")
    milestones_chart = milestone_chart(milestones, blue_line="ipdc_date")
    put_matplotlib_fig_into_word(doc, milestones_chart)

    return doc


wd_path = root_path / "input/summary_temp.docx"
report_doc = open_word_doc(wd_path)
m = Master(get_master_data(), get_project_information())

output = compile_report(report_doc, get_project_information(), m, Projects.crossrail)
output.save(root_path / "output/crossrail_report_tests.docx")


