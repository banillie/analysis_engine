from datamaps.api import project_data_from_master

from analysis_engine.data import (
    root_path,
    get_master_data,
    get_project_information,
    Master,
    JsonData,
)

# master = Master(get_master_data(), get_project_information())
path_str = str("{0}/core_data/json/master".format(root_path))
current = project_data_from_master(root_path / "core_data/master_3_2020.xlsx", 3, 2020)
JsonData(current, path_str)