"""
New production code for compiling total cost and benefits bar charts.
"""

from data_mgmt.data import Master, CostData, get_master_data, get_project_information, \
    Projects, total_costs_benefits_bar_chart_project, total_costs_benefits_bar_chart_group

master = Master(get_master_data(), get_project_information())
master.check_baselines()
costs = CostData(master)
costs.get_cost_totals_group('ipdc_costs')
total_costs_benefits_bar_chart_group(costs)
# costs.get_cost_totals_project(Projects.crossrail, 'ipdc_costs')
# total_costs_benefits_bar_chart_project(costs)
