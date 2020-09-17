#  Builds a vfm category type bar chart output in matplotlib

import sqlite3
#from vfm.database import convert_db_python_dict
from collections import Counter
import matplotlib.pyplot as plt
import numpy as np
from data_mgmt.data import root_path


# q_list = ['q1_2021', 'q4_1920']
# master_dict = convert_db_python_dict('vfm', q_list)


def query_db(db_name, key, quarter):
    conn = sqlite3.connect(db_name + '.db')
    conn.row_factory = lambda cursor, row: row[0]
    c = conn.cursor()
    c.execute("SELECT {key} FROM {table}".format(key=key, table=quarter))
    result = c.fetchall()

    conn.commit()
    conn.close()

    return result


def vfm_matplotlib_graph(labels, current_qrt, last_qrt):
    x = np.arange(len(labels))  # the label locations
    width = 0.35  # the width of the bars

    fig, ax = plt.subplots()
    rects_one = ax.bar(x - width / 2, current_qrt, width, label='This quarter')
    rects_two = ax.bar(x + width / 2, last_qrt, width, label='Last quarter')

    # Add some text for labels, title and custom x-axis tick labels, etc.
    ax.set_ylabel('Number')
    ax.set_title('Projects by VfM Category')
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.legend()

    def autolabel(rects):
        """Attach a text label above each bar in *rects*, displaying its height."""
        for rect in rects:
            height = rect.get_height()
            ax.annotate('{}'.format(height),
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 3),  # 3 points vertical offset
                        textcoords="offset points",
                        ha='center', va='bottom')

    autolabel(rects_one)
    autolabel(rects_two)

    fig.tight_layout()

    plt.show()

#  gets the number of projects reporting each category.
def get_vfm_cat_numbers(cat_list, qrt_dict):
    result = []
    for c in cat_list:
        try:
            result.append(qrt_dict[c])
        except KeyError:
            result.append(0)

    return result


ordered_cat_list = ['Poor', 'Low', 'Medium', 'High', 'Very High', 'Very High and Financially Positive']
cat_data_q1_2021 = query_db('vfm', 'vfm_cat_single', 'q1_2021')
cat_data_q4_1920 = query_db('vfm', 'vfm_cat_single', 'q4_1920')
q1_counter = Counter(cat_data_q1_2021)
q4_counter = Counter(cat_data_q4_1920)
q1_nos = get_vfm_cat_numbers(ordered_cat_list, q1_counter)
q4_nos = get_vfm_cat_numbers(ordered_cat_list, q4_counter)


vfm_matplotlib_graph(ordered_cat_list, q1_nos, q4_nos)