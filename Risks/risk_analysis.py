
from data_mgmt.data import (
    Master,
    root_path,
    get_project_information,
    get_master_data,
    RiskData, risks_into_excel,
)

m = Master(get_master_data(), get_project_information())
risk = RiskData(m)
wb = risks_into_excel(risk, "Q2 20/21")
wb.save(root_path / "output/risks.xlsx")