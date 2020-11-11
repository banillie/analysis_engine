from data_mgmt.data import Master, CostData, cost_profile_graph, \
    get_master_data, get_project_information

master = Master(get_master_data(), get_project_information())
master.check_baselines()
costs = CostData(master)
costs.get_cost_profile('ipdc_costs')
cost_profile_graph(costs, 'Group Test')
