"""
New code for compiling individual project reports.
"""

from datetime import date
from typing import Dict, Union, List

from docx import Document
from analysis_engine.data import (
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
    MilestoneData,
    milestone_chart,
    project_report_meta_data,
    change_word_doc_portrait,
    print_out_project_milestones,
    project_scope_text, make_file_friendly, string_conversion
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
    cost_profile = cost_profile_graph(costs, show="No")
    put_matplotlib_fig_into_word(doc, cost_profile)
    total_profile = total_costs_benefits_bar_chart(costs, benefits, show="No")
    put_matplotlib_fig_into_word(doc, total_profile)
    #  handling of no milestones within filtered period.
    ab = master.abbreviations[project_name]
    try:
        milestones.filter_chart_info(start_date="1/9/2020", end_date="30/12/2022")
        milestones_chart = milestone_chart(
            milestones, blue_line="ipdc_date", title=ab + " schedule (2021 - 22)", show="No"
        )
        put_matplotlib_fig_into_word(doc, milestones_chart)
        # print_out_project_milestones(doc, milestones, project_name)
    except ValueError:  # extends the time period.
        milestones = MilestoneData(master, project_name)
        milestones.filter_chart_info(start_date="1/9/2020", end_date="30/12/2024")
        milestones_chart = milestone_chart(
            milestones, blue_line="ipdc_date", title=ab + " schedule (2021 - 24)", show="No"
        )
        put_matplotlib_fig_into_word(doc, milestones_chart)
    print_out_project_milestones(doc, milestones, project_name)
    change_word_doc_portrait(doc)
    project_scope_text(doc, master, project_name)
    return doc


m = Master(get_master_data(), get_project_information())


def run_p_reports(projects: List[str] or str) -> None:
    projects = string_conversion(projects)
    for p in projects:
        report_doc = open_word_doc(root_path / "input/summary_temp.docx")
        qrt = make_file_friendly(str(m.master_data[0].quarter))
        output = compile_report(report_doc, get_project_information(), m, p)
        output.save(root_path / "output/{}_report_{}.docx".format(p, qrt))  # add quarter here


