
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
quarter_list = ["Q1 20/21", "Q4 19/20"]
wb = vfm_into_excel(vfm, quarter_list)
wb.save(root_path / "output/vfm.xlsx")
