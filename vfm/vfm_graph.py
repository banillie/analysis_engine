#  Builds a vfm category type bar chart output in matplotlib

from collections import Counter
import matplotlib.pyplot as plt
import numpy as np
from data_mgmt.data import root_path, vfm_matplotlib_graph
from vfm.database import query_db

#  gets the number of projects reporting each category.
def get_vfm_cat_numbers(cat_list, qrt_dict):
    result = []
    for c in cat_list:
        try:
            result.append(qrt_dict[c])
        except KeyError:
            result.append(0)

    return result


ordered_cat_list = ['Poor', 'Low', 'Medium', 'High', 'Very High',
                    'Very High and Financially Positive', 'Economically Positive']
db_path = root_path / "core_data/vfm.db"
cat_data_q1_2021 = query_db(db_path, 'vfm_cat_single', 'q1_2021')
cat_data_q4_1920 = query_db(db_path, 'vfm_cat_single', 'q4_1920')
q1_counter = Counter(cat_data_q1_2021)
q4_counter = Counter(cat_data_q4_1920)
q1_nos = get_vfm_cat_numbers(ordered_cat_list, q1_counter)
q4_nos = get_vfm_cat_numbers(ordered_cat_list, q4_counter)

vfm_matplotlib_graph(ordered_cat_list, q1_nos, q4_nos, 'test')