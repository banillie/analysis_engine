from analysis_engine.data import open_pickle_file, root_path, CostData, MilestoneData, cost_v_schedule_chart_into_wb

# print("compiling cost and schedule matrix analysis")
m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
costs = CostData(m, m.current_projects)
miles = MilestoneData(m, m.current_projects)
miles.calculate_schedule_changes()
wb = cost_v_schedule_chart_into_wb(miles, costs)
wb.save(root_path / "output/costs_schedule_matrix.xlsx")
# print("Cost and schedule matrix compiled. Enjoy!")
