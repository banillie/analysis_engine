"""
outputs all sro confidence ratings. outputs are:
- word document printout of which project dca ratings have changed
"""

from data_mgmt.data import (
    Master,
    root_path,
    calculate_dca_change,
    dca_changes_into_word,
    DCA_KEYS,
    get_project_information,
    get_master_data,
    get_word_doc
)


def dca_analysis(m_data, project_info, word_doc):
    m = Master(m_data, project_info)
    assessment = calculate_dca_change(m, DCA_KEYS["SRO"])
    dca_changes_into_word(assessment, word_doc)
    word_doc.save(root_path / "output/dca_checks.docx")


dca_analysis(get_master_data(), get_project_information(), get_word_doc())
