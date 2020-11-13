"""
New production code for compiling total cost and benefits bar charts.
"""

from data_mgmt.data import Master, CostData, get_master_data, get_project_information, \
    Projects, total_costs_benefits_bar_chart, cost_profile_graph, BenefitsData, cost_profile_baseline_graph

master = Master(get_master_data(), get_project_information())
master.check_baselines()
costs = CostData(master)
benefits = BenefitsData(master)


def interactive_totals(project, *args):
    costs.get_cost_totals(project, 'ipdc_costs')
    benefits.get_ben_totals(project, 'ipdc_benefits')
    if args == ():
        total_costs_benefits_bar_chart(costs, benefits)
    else:
        total_costs_benefits_bar_chart(costs, benefits, args[0])


def interactive_standard_profile(project, *args):
    costs.get_cost_profile(project, 'ipdc_costs')
    if args == ():
        cost_profile_graph(costs)
    else:
        cost_profile_graph(costs, args[0])


def interactive_baseline_profile(project, *args):
    costs.get_cost_profile(project, 'ipdc_costs')
    if args == ():
        cost_profile_baseline_graph(costs)
    else:
        cost_profile_baseline_graph(costs, args[0])

