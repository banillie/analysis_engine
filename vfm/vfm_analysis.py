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


m = Master(get_master_data(), get_project_information())
vfm = VfMData(m)
quarter_list = ["Q2 20/21", "Q1 20/21"]
wb = vfm_into_excel(vfm, quarter_list)
wb.save(root_path / "output/vfm.xlsx")
