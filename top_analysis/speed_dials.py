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
    get_word_doc, DcaData, dca_changes_into_excel
)


m = Master(get_master_data(), get_project_information())
dca = DcaData(m)
# dca.get_changes("Q2 20/21", "Q1 20/21")
# word_doc = dca_changes_into_word(dca, get_word_doc())
# word_doc.save(root_path / "output/dca_checks.docx")
quarter_list = ["Q2 20/21", "Q1 20/21"]
wb = dca_changes_into_excel(dca, quarter_list)
wb.save(root_path / "output/dca_print.xlsx")


# dca_analysis(get_master_data(), get_project_information(), get_word_doc())
