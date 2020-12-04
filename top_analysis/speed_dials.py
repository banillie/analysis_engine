"""
outputs all analysis for sro confidence ratings. outputs are:
- word document printout of which project dca ratings have changed
"""

from data_mgmt.data import (
    Master,
    root_path,
    dca_changes_into_word,
    get_project_information,
    get_master_data,
    get_word_doc, DcaData
)


def dca_analysis(m_data, project_info, word_doc):
    m = Master(m_data, project_info)
    dca = DcaData(m, m.current_projects)
    dca.get_changes("Q2 20/21", "Q1 20/21")
    dca_changes_into_word(dca, word_doc)
    word_doc.save(root_path / "output/dca_checks.docx")


dca_analysis(get_master_data(), get_project_information(), get_word_doc())
