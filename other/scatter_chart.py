from analysis_engine.data import open_pickle_file, root_path, DcaData, CostData, MilestoneData, \
    cost_v_schedule_chart_into_wb, RiskData, DandelionData, VfMData

# print("compiling cost and schedule matrix analysis")
m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
# dandelion = DandelionData(m, quarter="standard")
# costs = CostData(m, baseline='all')
vfm = VfMData(m, quarter=['standard'])
# dca = DcaData(m)
# dca.get_changes()
# risk = RiskData(m, group=["Rail"])
# miles = MilestoneData(m, baseline='all')
# miles.calculate_schedule_changes()
# wb = cost_v_schedule_chart_into_wb(miles, costs)
# wb.save(root_path / "output/costs_schedule_matrix.xlsx")
# # print("Cost and schedule matrix compiled. Enjoy!")
