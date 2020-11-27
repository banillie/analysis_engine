"""
New production code for compiling total cost and benefits bar charts.
"""

from data_mgmt.data import Master, get_master_data, get_project_information, root_path, CostData, \
    BenefitsData, totals_chart, Projects, FIGURE_STYLE, standard_profile

master = Master(get_master_data(), get_project_information())
# master.check_baselines()
costs = CostData(master)
benefits = BenefitsData(master)
wd_path = root_path / "input/summary_temp_landscape.docx"
fig_style_1 = FIGURE_STYLE[1]  # half horizontal
fig_style_2 = FIGURE_STYLE[2]  # full horizontal





