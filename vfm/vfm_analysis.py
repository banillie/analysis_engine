"""
Compiles VfM analysis. Places into output file an excel file with required data.
User can specify which quarters data they would like to return in the quarters_list variable.
"""

from data_mgmt.data import (
    Master,
    root_path,
    get_project_information,
    get_master_data,
    VfMData,
    vfm_into_excel
)


def compile_vfm_analysis():
    m = Master(get_master_data(), get_project_information())
    vfm = VfMData(m)
    latest_quarter = str(m.master_data[0].quarter)
    last_quarter = str(m.master_data[1].quarter)
    default_quarter_list = [latest_quarter, last_quarter]
    wb = vfm_into_excel(vfm, default_quarter_list)
    wb.save(root_path / "output/vfm.xlsx")


compile_vfm_analysis()
