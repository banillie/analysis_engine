import numpy as np

from analysis_engine.data import open_pickle_file, root_path, DcaData, CostData, MilestoneData, \
    cost_v_schedule_chart_into_wb, RiskData, DandelionData, VfMData, put_cost_totals_into_wb, \
    put_matplotlib_fig_into_word, open_word_doc, FIGURE_STYLE, get_cost_stackplot_data, put_stackplot_data_into_wb, \
    make_a_dandelion_auto

# print("compiling cost and schedule matrix analysis")
m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
# dandelion = DandelionData(m, quarter="standard")
# d = {}
d_data = DandelionData(m)
dlion = make_a_dandelion_auto(d_data)

# g = m.dft_groups["Q3 20/21"]["Rail"]
# group_list = ['HSMRPG', 'Rail', 'RPE', 'AMIS']
# c = get_cost_stackplot_data(m, group_list, "Q3 20/21", type="comp")
# put_stackplot_data_into_wb(c)
# doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# put_matplotlib_fig_into_word(doc, c, size=8, transparent=False)
# doc.save(root_path /"output/portfolio_cost_comp_cat_no_rf_npr.docx")
# costs = CostData(m, quarter=['standard'])

# wb = put_cost_totals_into_wb(costs)
# wb.save(root_path / "output/cost_totals.xlsx")
# vfm = VfMData(m, quarter=['standard'])
# dca = DcaData(m)
# dca.get_changes()
# risk = RiskData(m, quarter=["standard"])
# miles = MilestoneData(m, baseline='all')
# miles.calculate_schedule_changes()
# wb = cost_v_schedule_chart_into_wb(miles, costs)
# wb.save(root_path / "output/costs_schedule_matrix.xlsx")
# # print("Cost and schedule matrix compiled. Enjoy!")


# def moving_average(x, w):
#     return np.convolve(x, np.ones(w), 'valid') / w

