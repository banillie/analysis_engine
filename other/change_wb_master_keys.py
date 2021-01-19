"""
This code changes keys names in master wb documents
"""

from analysis_engine.data import get_master_data_file_paths, root_path, put_key_change_master_into_dict, \
    run_change_keys, get_master_data_file_paths_fy_16_17, get_master_data_file_paths_fy_19_20, \
    get_master_data_file_paths_fy_18_19, get_master_data_file_paths_fy_17_18

keys_dict = put_key_change_master_into_dict(root_path / "core_data/analysis_engine/keys_to_change.xlsx")
run_change_keys(get_master_data_file_paths_fy_16_17(), keys_dict)
