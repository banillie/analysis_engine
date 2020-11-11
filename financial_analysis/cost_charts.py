"""
New production code for compiling total cost and benefits bar charts.
"""

from data_mgmt.data import Master, CostData, get_master_data, get_project_information, \
    Projects, total_costs_benefits_bar_chart_project, total_costs_benefits_bar_chart, \
    cost_profile_graph, BenefitsData, project_cost_profile_graph

master = Master(get_master_data(), get_project_information())
master.check_baselines()
costs = CostData(master)
benefits = BenefitsData(master)

# GROUPS
costs.get_cost_totals(master.current_projects, 'ipdc_costs')
benefits.get_ben_totals(master.current_projects, 'ipdc_benefits')
total_costs_benefits_bar_chart(costs, benefits, 'IPDC Portfolio')

# PROJECTS
# costs.get_cost_totals_project(Projects.a14, 'ipdc_costs')
# benefits.get_ben_totals_project(Projects.a14, 'ipdc_benefits')
# total_costs_benefits_bar_chart_project(costs, benefits)
