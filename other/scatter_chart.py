import math
import numpy as np
# from typing import List
from matplotlib import pyplot as plt

from analysis_engine.data import (
    open_pickle_file,
    root_path,
    # DcaData,
    # CostData,
    # MilestoneData,
    # put_milestones_into_wb, data_query_into_wb, milestone_chart,
    # cost_v_schedule_chart_into_wb,
    # RiskData,
    DandelionData,
    # VfMData,
    # put_cost_totals_into_wb,
    put_matplotlib_fig_into_word,
    open_word_doc,
    # FIGURE_STYLE,
    # get_cost_stackplot_data,
    # put_stackplot_data_into_wb,
    make_a_dandelion_auto,
    # BenefitsData,
    # total_costs_benefits_bar_chart,
    # gauge,
    calculate_arg_combinations, get_sp_data, cal_group, cost_stackplot_graph, get_sp_data,
    # get_master_data,
    # get_project_information,
    # Master
)

## GENERATE CLI OPTIONS
# arg_list = ["quarters", "group", "stage", "remove", "type"]
# calculate_arg_combinations(arg_list)

## INITIATE
# master = Master(get_master_data(), get_project_information())

## MASTER
m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))

## DANDELION
# stage = ["pre-SOBC", "SOBC", "OBC", "FBC"]
# group = ["HSRG", "AMIS", "RIG", "RSS", "RPE"]
# d_data = DandelionData(m, quarter=[str(m.current_quarter)], group=group, meta="benefits")
# d_lion = make_a_dandelion_auto(d_data, chart=True)
# doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# put_matplotlib_fig_into_word(doc, d_lion, size=7.5)
# doc.save(root_path / "output/dlion_graph.docx")

## MILESTONES
# ms = MilestoneData(m, quarter=[str(m.current_quarter)], group=["GMPP"])
# ms.filter_chart_info(dates=["1/1/2021", "1/1/2030"])
# wb = put_milestones_into_wb(ms)
# wb.save(root_path / "output/gmpp_milestones_data.xlsx")
# doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# for p in m.dft_groups["Q3 20/21"]["GMPP"]:
#     try:
#         ms = MilestoneData(m, quarter=[str(m.current_quarter)], group=[p])
#         ms.filter_chart_info(dates=["1/1/2021", "1/1/2030"])
#         graph = milestone_chart(ms, blue_line="Today", chart=True)
#         put_matplotlib_fig_into_word(doc, graph, size=8, transparent=False)
#     except ValueError:
#         pass
# doc.save(root_path / "output/gmpp_milestones_charts.docx".format(p))

# b = BenefitsData(m, baseline=["all"])
# total_costs_benefits_bar_chart(c, b, chart=True)

# QUERY
# DFT_GROUP = ["HSMRPG", "AMIS", "Rail", "RPE"]
# wb = data_query_into_wb(m, keys=["Senior Responsible Owner (SRO)"], quarter=["Q3 19/20"], group=DFT_GROUP)
# wb.save(root_path / "output/query_test.xlsx")

## STACKPLOT
DFT_GROUP = ["HSRG", "RSS", "RIG", "AMIS", "RPE"]
# DFT_GROUP = ["RPE", "AMIS"]
# g = cal_group(["FBC"], m, 0)
# sp_data = get_sp_data(m, g, [str(m.current_quarter)], type="comp", remove=["HS2 1"])
sp_data = get_sp_data(m, stage=["SOBC"], quarter=["standard"], type="cat")
cost_stackplot_graph(sp_data)
