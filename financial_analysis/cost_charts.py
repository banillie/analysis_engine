"""
New production code for compiling total cost and benefits bar charts.
"""

from data_mgmt.data import Master, CostData, get_master_data, get_project_information, \
    Projects, total_costs_benefits_bar_chart_project, total_costs_benefits_bar_chart_group, \
    group_cost_profile_graph, BenefitsData

master = Master(get_master_data(), get_project_information())
master.check_baselines()
costs = CostData(master)
benefits = BenefitsData(master)

# GROUPS
# costs.get_cost_totals_group(Projects.fbc_stage, 'ipdc_costs')
# benefits.get_ben_totals_group(Projects.fbc_stage, 'ipdc_benefits')
# total_costs_benefits_bar_chart_group(costs, benefits, 'Total FBC Group')

# PROJECTS
costs.get_cost_totals_project(Projects.crossrail, 'ipdc_costs')
benefits.get_ben_totals_project(Projects.crossrail, 'ipdc_benefits')
total_costs_benefits_bar_chart_project(costs, benefits)
