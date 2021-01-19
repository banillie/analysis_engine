#  code of running and compiling vfm data analysis each quarter.

from analysis.data import root_path, get_master_data
from vfm.vfm_analysis_workings_old import compile_data


master_data = get_master_data()
current_project_name_list = master_data[0].projects
ordered_cat_list = ['Poor', 'Low', 'Medium', 'High', 'Very High',
                    'Very High and Financially Positive', 'Economically Positive',
                    None]
run = compile_data(master_data, current_project_name_list, ordered_cat_list)
run.save(root_path / "output/vfm_data_output.xlsx")