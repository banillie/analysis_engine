#  code of running and compiling vfm data analysis each quarter.

from data_mgmt.data import root_path
from database.database import convert_db_python_dict, get_project_names
from vfm.vfm_analysis_workings import compile_data_db

q_list = ['q1_2021', 'q4_1920']
db_path = root_path / "core_data/vfm.db"
master_dict = convert_db_python_dict(db_path, q_list)
project_names = get_project_names(db_path, 'q1_2021')
ordered_cat_list = ['Poor', 'Low', 'Medium', 'High', 'Very High',
                    'Very High and Financially Positive', 'Economically Positive',
                    None]
run = compile_data_db(master_dict, project_names, ordered_cat_list)
run.save(root_path / "output/vfm_data_wb_output.xlsx")