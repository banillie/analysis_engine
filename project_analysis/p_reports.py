"""
New code for compiling individual project reports.
"""
import os
from datetime import date
from typing import Dict, Union

from docx import Document
from data_mgmt.data import root_path, SRO_conf_key_list, project_cost_profile_graph, CostData, \
    get_project_information, get_master_data, Master, current_projects, open_word_doc, wd_heading, key_contacts, \
    dca_table, dca_narratives, year_cost_profile_chart


def compile_report(doc: Document, project_info: Dict[str, Union[str, int, date, float]], master, project) -> Document:
    wd_heading(doc, project_info, project)
    key_contacts(doc, master, project)
    dca_table(doc, master, project)
    dca_narratives(doc, master, project)
    costs = CostData(master)
    costs.get_profile_project(project, 'ipdc_costs')
    year_cost_profile_chart(doc, costs)
    return doc


wd_path = root_path / "input/summary_temp.docx"
report_doc = open_word_doc(wd_path)
project_information = get_project_information()
live_projects = current_projects(project_information)
master_data = get_master_data()
m = Master(master_data, ["Crossrail Programme"])

output = compile_report(report_doc, project_information, m, "Crossrail Programme")
output.save(root_path / "output/crossrail_report_test.docx")