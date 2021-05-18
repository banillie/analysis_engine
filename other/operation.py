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
    milestone_chart,
    run_p_reports,
    put_milestones_into_wb,
    cost_v_schedule_chart_into_wb,
    RiskData,
    DandelionData,
    VfMData,
    vfm_into_excel,
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
    risks_into_excel,
    build_speedials,
    dca_changes_into_word,
    radar_chart,
    JsonData,
    open_json_file, get_input_doc, ipdc_dashboard,
)

## GENERATE CLI OPTIONS
# arg_list = ["quarters", "group", "stage", "remove", "type"]
# calculate_arg_combinations(arg_list)

## INITIATE
from analysis_engine.top35_data import (
    top35_get_master_data,
    top35_get_project_information,
    top35_root_path, top35_run_p_reports,
)

# master = Master(get_master_data(), get_project_information())

## PICKLE
m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))

## GROUPS
DFT_GROUP = ["HSRG", "RSS", "RIG", "AMIS", "RPE"]
STAGE_GROUPS = ["pipeline", "pre-SOBC", "SOBC", "OBC", "FBC"]

word_doc_landscape = open_word_doc(root_path / "input/summary_temp_landscape.docx")


top35_data_dict = {
    "docx_save_path": str(top35_root_path / "output/{}.docx"),
    "master": Master(top35_get_master_data(), top35_get_project_information(), data_type="top35"),
    "op_args": {
        "quarter": ["Q4 20/21"],
        "group": ["HSRG", "RSS", "RIG", "RPE"],
        "chart": False,
        "data_type": "top35",
        "circle_colour": "No",
    },
}

ipdc_data_dict = {
    "docx_save_path": str(root_path / "output/{}.docx"),
    "master": open_pickle_file(str(root_path / "core_data/pickle/master.pickle")),
    "op_args": {
        "quarter": ["Q4 20/21"],
        "baseline": ["standard"],
        # "group": ["HSRG", "RSS", "RIG", "AMIS", "RPE"],
        "group": ["HS2 2a"],
        "chart": True,
        "circle_colour": "No",
    },
}

# cdg_data_dict = {
#     "docx_save_path": cdg_root_path / "output/{}.docx",
#     "data": (cdg_get_master_data(), cdg_get_project_information()),
#     "op_args": {
#         "quarter": ["Q4 20/21"],
#         "group": ["GF", "CFPD", "SCS"],
#         "chart": True,
#         "data_type": "cdg",
#     },
# }

data = ipdc_data_dict

## DANDELION
# data = top35_data_dict
# dl_data = DandelionData(data["master"], **data["op_args"])
# d_lion = make_a_dandelion_auto(dl_data, **data["op_args"])
# put_matplotlib_fig_into_word(word_doc_landscape, d_lion, size=7)
# word_doc_landscape.save(data["docx_save_path"].format("dandelion_graph"))


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

## SUMMARIES
# run_p_reports(data["master"], **data["op_args"])

## VFM
# c = VfMData(m, group=DFm = Master(*data["data"], **data["op_args"] )T_GROUP, quarter=["standard"])  # c is class
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
# t = m.project_stage
# for s in t["Q4 20/21"].keys():
#     g = t["Q4 20/21"][s]
#     if not g:
#         continue
# # for g in m.current_projects:
#     else:
#         chart = radar_chart(sp_data, m, group=g)
#         doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
#         put_matplotlib_fig_into_word(doc, chart, size=5)
#         doc.save(root_path / "output/{}_radar_5_poly_individual.docx".format(s))
# chart = radar_chart(sp_data, m, title="All")
# put_matplotlib_fig_into_word(doc, chart, size=5)
# doc.save(root_path / "output/radar_5_poly_all.docx")


## DASHBOARD
dashboard_master = get_input_doc(root_path / "input/dashboards_master.xlsx")
wb = ipdc_dashboard(m, dashboard_master)
wb.save(root_path / "output/completed_ipdc_dashboard.xlsx")
