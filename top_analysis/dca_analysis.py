"""
Outputs analysis for SRO confidence ratings. Output is placed in analysis_engine/output. The output is:
- An excel workbook titled dca_data which provides a count of DCAs and their proportion of cost.
"""

from analysis_engine.data import (
    Master,
    root_path,
    get_project_information,
    get_master_data,
    DcaData,
    dca_changes_into_excel,
)


def compile_dca_analysis():
    m = Master(get_master_data(), get_project_information())
    dca = DcaData(m)
    latest_quarter = str(m.master_data[0].quarter)
    last_quarter = str(m.master_data[1].quarter)
    default_quarter_list = [latest_quarter, last_quarter]
    wb = dca_changes_into_excel(dca, default_quarter_list)
    wb.save(root_path / "output/dca_data.xlsx")


compile_dca_analysis()
