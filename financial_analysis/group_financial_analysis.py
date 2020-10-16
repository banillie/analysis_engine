from data_mgmt.data import Master, CostData, group_cost_profile_graph, current_projects, project_information, \
    get_master_data

live_projects = current_projects(project_information)
master_data = get_master_data()
master = Master(master_data, live_projects)
costs = CostData(master)
costs.get_profile_all('ipdc_costs')
group_cost_profile_graph(costs, 'Group Test')
