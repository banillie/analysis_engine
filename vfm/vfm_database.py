#  code for creating/connecting to vfm database and adding data.

from database.database import get_vfm_values, get_quarter_values, \
    create_vfm_table, insert_many_vfm_db
from data_mgmt.data import get_master_data

m = get_master_data()
vfm_db_list = get_vfm_values(m)
# vfm_q1_2021 = get_quarter_values(vfm_db_list, "Q1 20/21")
# vfm_q4_1920 = get_quarter_values(vfm_db_list, "Q4 19/20")
# vfm_q3_1920 = get_quarter_values(vfm_db_list, "Q3 19/20")
vfm_q2_1920 = get_quarter_values(vfm_db_list, "Q2 19/20")

create_vfm_table('vfm', 'q2_1920')
insert_many_vfm_db('vfm', 'q2_1920', vfm_q2_1920)



