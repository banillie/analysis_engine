#  code of running and compiling vfm data analysis each quarter.

from data_mgmt.data import root_path, get_master_data, get_current_project_names
from vfm.vfm_analysis_workings import compile_data


master_data = get_master_data()
current_project_name_list = get_current_project_names()
ordered_cat_list = ['Poor', 'Low', 'Medium', 'High', 'Very High',
                    'Very High and Financially Positive', 'Economically Positive',
                    None]
run = compile_data(master_data, current_project_name_list, ordered_cat_list)
run.save(root_path / "output/vfm_data_output.xlsx")