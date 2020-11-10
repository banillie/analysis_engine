"""
New production code for compiling total cost and benefits bar charts.
"""

from data_mgmt.data import Master, CostData, get_master_data, get_project_information, \
    Projects, total_costs_benefits_bar_chart_project, total_costs_benefits_bar_chart_group, \
    group_cost_profile_graph, BenefitsData

master = Master(get_master_data(), get_project_information())
master.check_baselines()
costs = CostData(master)
bens = BenefitsData(master)
costs.get_cost_totals_group('ipdc_costs')
bens.get_ben_totals_group('ipdc_benefits')
total_costs_benefits_bar_chart_group(costs, bens, 'Total Group')

# costs.get_cost_totals_group('ipdc_costs')
# total_costs_benefits_bar_chart_group(costs)
# costs.get_cost_totals_project(Projects.crossrail, 'ipdc_costs')
# total_costs_benefits_bar_chart_project(costs)
# costs.get_profile_group('ipdc_costs')
# group_cost_profile_graph(costs, 'Portfolio')