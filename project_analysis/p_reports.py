"""
New code for compiling individual project reports.
"""
import os
from datetime import date
from typing import Dict, Union

from docx import Document
from data_mgmt.data import root_path, CostData, Projects,\
    get_master_data, Master, open_word_doc, wd_heading, key_contacts, get_project_information, \
    dca_table, dca_narratives, put_matplotlib_fig_into_word


def compile_report(doc: Document, project_info: Dict[str, Union[str, int, date, float]], master: Master, project_name: str) -> Document:
    wd_heading(doc, project_info, project_name)
    key_contacts(doc, master, project_name)
    dca_table(doc, master, project_name)
    dca_narratives(doc, master, project_name)
    costs = CostData(master)
    costs.get_profile_project(project_name, 'ipdc_costs')
    put_matplotlib_fig_into_word(doc, costs)
    return doc


wd_path = root_path / "input/summary_temp.docx"
report_doc = open_word_doc(wd_path)
m = Master(get_master_data(), get_project_information())

output = compile_report(report_doc, get_project_information(), m, Projects.crossrail)
output.save(root_path / "output/crossrail_report_tests.docx")