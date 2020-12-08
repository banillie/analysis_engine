"""
Outputs analysis for SRO confidence ratings. Outputs are place in analysis_engine/output. They are:
- A word document titled dca_changes which specifies which project dca ratings have changed
- An excel workbook titled dca_data which provides a count of DCAs and their proportion of cost.
"""

from data_mgmt.data import (
    Master,
    root_path,
    dca_changes_into_word,
    get_project_information,
    get_master_data,
    get_word_doc,
    DcaData,
    dca_changes_into_excel,
)


def compile_dca_analysis():
    m = Master(get_master_data(), get_project_information())
    dca = DcaData(m)
    dca.get_changes("Q2 20/21", "Q1 20/21")
    word_doc = dca_changes_into_word(dca, get_word_doc())
    word_doc.save(root_path / "output/dca_changes.docx")
    quarter_list = ["Q2 20/21", "Q1 20/21"]
    wb = dca_changes_into_excel(dca, quarter_list)
    wb.save(root_path / "output/dca_data.xlsx")


compile_dca_analysis()
