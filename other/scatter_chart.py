import json
import math
import numpy as np
# from typing import List
from matplotlib import pyplot as plt

# matplotlib==3.3.0

from analysis_engine.data import (
    open_pickle_file,
    root_path,
    DcaData,
    CostData,
    MilestoneData,
    put_milestones_into_wb,
    data_query_into_wb,
    milestone_chart, run_p_reports, put_milestones_into_wb,
    cost_v_schedule_chart_into_wb,
    RiskData,
    DandelionData,
    VfMData, vfm_into_excel,
    put_cost_totals_into_wb,
    put_matplotlib_fig_into_word,
    open_word_doc,
    FIGURE_STYLE,
    put_stackplot_data_into_wb,
    make_a_dandelion_auto,
    BenefitsData,
    total_costs_benefits_bar_chart,
    gauge,
    calculate_arg_combinations,
    get_sp_data,
    cal_group,
    cost_stackplot_graph,
    get_sp_data,
    cost_profile_graph,
    get_master_data,
    get_project_information,
    Master,
    risks_into_excel, build_speedials, dca_changes_into_word, radar_chart, JsonData, open_json_file,
)

## GENERATE CLI OPTIONS
# arg_list = ["quarters", "group", "stage", "remove", "type"]
# calculate_arg_combinations(arg_list)

## INITIATE
# master = Master(get_master_data(), get_project_information())


## JSON
# master = Master(get_master_data(), get_project_information())
# json_m = JsonData(master, "{}/core_data/json/master".format(root_path))
m = open_json_file(str(root_path / "core_data/json/master.json"))


## PICKLE
# m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))

## GROUPS
DFT_GROUP = ["HSRG", "RSS", "RIG", "AMIS", "RPE"]
STAGE_GROUPS = ["pipeline", "pre-SOBC", "SOBC", "OBC", "FBC"]

doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")


## OP_ARGS
# op_args = {
#     # "quarter": ["standard"],
#     "quarter": [str(m.current_quarter)],
#     # "group": DFT_GROUP,
#     "stage": ["pipeline", "pre-SOBC", "SOBC", "OBC", "FBC"],
#     # "stage": "Ely Area Capacity Enhancement Programme",
#     "type": "remaining",
#     # "chart": True,
#     # "baseline": ["standard"],
#     # "remove": "HS2 1",
#     # "confidence": "benefits"
#     "angles": [260, 300, 345, 40]
#     }

# ## DANDELION
# stage = ["pre-SOBC", "SOBC", "OBC", "FBC"]
# d_data = DandelionData(m, **op_args)
# d_lion = make_a_dandelion_auto(d_data, **op_args)
# put_matplotlib_fig_into_word(doc, d_lion, size=7.5)
# doc.save(root_path / "output/dlion_graph.docx")
#

## MILESTONES
# ms = MilestoneData(m, **op_args)
# ms.filter_chart_info(dates=["1/4/2021", "1/5/2021"])
# wb = put_milestones_into_wb(ms)
# wb.save(root_path / "output/test_milestone_data_output.xlsx")
# chart_kwargs = {**{"blue_line": "today", "Chart": True}, **ms.kwargs}
# milestone_chart(ms, m, **chart_kwargs)
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

# # QUERY
# # wb = data_query_into_wb(m, keys=["Senior Responsible Owner (SRO)"], quarter=["Q3 19/20"], group=DFT_GROUP)
# # wb.save(root_path / "output/query_test.xlsx")

# ## STACKPLOT
# sp_data = get_sp_data(m, quarter=["standard"], group=["RIG"])
# cost_stackplot_graph(sp_data, m, group=DFT_GROUP)

# COSTS
# c = CostData(m, **op_args)
# cost_profile_graph(c, m, chart=True, group=c.start_group)

## SUMMARYS
# run_p_reports(m, **op_args)

## VFM
# c = VfMData(m, group=DFT_GROUP, quarter=["standard"])  # c is class
# wb = vfm_into_excel(c)

## RISKS
# c = RiskData(m, **op_args)
# wb = risks_into_excel(c)

## SPEED DIALS
# data = DcaData(m, **op_args)
# data.get_changes()
# land_doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# build_speedials(data, land_doc)
# land_doc.save(root_path / "output/speed_dial_graph.docx")
# doc = open_word_doc(root_path / "input/summary_temp.docx")
# doc = dca_changes_into_word(data, doc)
# doc.save(root_path / "output/speed_dials.docx")


## RADAR CHART
# sp_data = root_path / "core_data/sp_master.xlsx"
# t = m.dft_groups
# for s in t["Q4 20/21"].keys():
#     g = t["Q4 20/21"][s]
#     if not g:
#         continue
# for g in m.current_projects:
#     chart = radar_chart(sp_data, m, group=[g], title=m.abbreviations[g]["abb"])
#     put_matplotlib_fig_into_word(doc, chart, size=5)
#     doc.save(root_path / "output/radar_individual.docx")
# chart = radar_chart(sp_data, m, title="All")
# put_matplotlib_fig_into_word(doc, chart, size=5)
# doc.save(root_path / "output/radar_all.docx")