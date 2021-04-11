from analysis_engine.data import (
    root_path,
    get_master_data,
    project_data_from_master,
    get_project_information,
    Master,
    Json,
)

# master = Master(get_master_data(), get_project_information())
path_str = str("{0}/core_data/json/master".format(root_path))
current = project_data_from_master(root_path / "core_data/master_3_2020.xlsx", 4, 2020)
Json(current, path_str)