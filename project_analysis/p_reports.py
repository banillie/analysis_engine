"""
New code for compiling individual project reports.
"""
import os
from datetime import date
from typing import Dict, Union

from docx import Document
from data_mgmt.data import root_path, SRO_CONF_KEY_LIST, project_cost_profile_graph, CostData, \
    get_master_data, Master, open_word_doc, wd_heading, key_contacts, get_project_information, \
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
m = Master(get_master_data(), get_project_information())

output = compile_report(report_doc, get_project_information(), m, "East West Rail Configuration State 1")
output.save(root_path / "output/ewr_cs1_report_test.docx")