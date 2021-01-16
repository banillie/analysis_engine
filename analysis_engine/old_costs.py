"""
Code used for getting old financial year data that should be present in current masters.
"""

from analysis_engine.data import root_path, get_master_data_file_paths_fy_18_19, run_place_old_fy_data_into_masters

# get_old_fy_cost_data(root_path / "core_data/master_4_2018.xlsx",
#                      root_path / "core_data/other/project_info_fy_18_19_cost_info.xlsx")

run_place_old_fy_data_into_masters(get_master_data_file_paths_fy_18_19(),
                                   root_path / "core_data/other/project_info_fy_17_18_cost_info.xlsx")
