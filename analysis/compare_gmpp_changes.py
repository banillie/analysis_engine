from analysis.data import get_gmpp_projects, compare_masters, get_project_information, \
    get_master_data_file_paths_fy_20_21, root_path

gmpp_list = get_gmpp_projects(get_project_information())
wb = compare_masters(get_master_data_file_paths_fy_20_21(), gmpp_list)
wb.save(root_path / "output/gmpp_data_compared.xlsx")
