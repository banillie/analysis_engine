# #  code for creating/connecting to vfm database and adding data.
#
# from database.database import get_vfm_values, get_quarter_values, \
#     create_vfm_table, insert_many_vfm_db
# from analysis.data import get_master_data
#
# m = get_master_data()
# vfm_db_list = get_vfm_values(m)
# # vfm_q1_2021 = get_quarter_values(vfm_db_list, "Q1 20/21")
# # vfm_q4_1920 = get_quarter_values(vfm_db_list, "Q4 19/20")
# # vfm_q3_1920 = get_quarter_values(vfm_db_list, "Q3 19/20")
# vfm_q2_1920 = get_quarter_values(vfm_db_list, "Q2 19/20")
#
# create_vfm_table('vfm', 'q2_1920')
# insert_many_vfm_db('vfm', 'q2_1920', vfm_q2_1920)
#


from analysis.data import get_master_data, root_path
from database.database import create_db, import_master_to_db
import os
from datamaps.api import project_data_from_master

db_path = os.path.join(os.getcwd(), "live.db")
create_db(db_path)

all_m = get_master_data()
groups_ids = project_data_from_master(root_path / 'core_data/project_group_id_no.xlsx', 1, 2099)

import_master_to_db(db_path, all_m, groups_ids)