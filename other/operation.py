import json
import math
import numpy as np

# from typing import List
from matplotlib import pyplot as plt

# matplotlib==3.3.0
from analysis_engine.cdg_data import (
    cdg_root_path,
    cdg_get_master_data,
    cdg_get_project_information,
    cdg_dashboard,
    cdg_run_p_reports, cdg_narrative_dashboard,
)
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
    open_json_file,
    get_input_doc,
    ipdc_dashboard,
)

## GENERATE CLI OPTIONS
# arg_list = ["quarters", "group", "stage", "remove", "type"]
# calculate_arg_combinations(arg_list)

## INITIATE
from analysis_engine.top35_data import (
    top35_get_master_data,
    top35_get_project_information,
    top35_root_path,
    top35_run_p_reports,
)

# master = Master(get_master_data(), get_project_information())

## PICKLE
# m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))

## GROUPS
DFT_GROUP = ["HSRG", "RSS", "RIG", "AMIS", "RPE"]
STAGE_GROUPS = ["pipeline", "pre-SOBC", "SOBC", "OBC", "FBC"]
hoz_doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
doc = open_word_doc(root_path / "input/summary_temp.docx")

# top35_data_dict = {
#     "docx_save_path": str(top35_root_path / "output/{}.docx"),
#     "master": Master(top35_get_master_data(), top35_get_project_information(), data_type="top35"),
#     "op_args": {
#         "quarter": ["Month(May), 2021"],
#         "group": ["HSRG", "RSS", "RIG", "RPE"],
#         # "group": ["HS2 Prog"],
#         "chart": False,
#         "data_type": "top35",
#         "circle_colour": "No",
#     },
#     "excel_save_path": str(top35_root_path / "output/{}.xlsx"),
#     "word_save_path": str(top35_root_path / "output/{}.docx")
# }

ipdc_data_dict = {
    "docx_save_path": str(root_path / "output/{}.docx"),
    "master": Master(open_json_file(str(root_path / "core_data/json/master.json"))),
    "op_args": {
        "quarter": ["standard"],
        # "quarter": ["Q4 20/21"],
        # "baseline": ["bl_one"],
        # "group": ["HSRG", "RSS", "RIG", "AMIS", "RPE"],
        "group": ["RIG"],
        "dates": ["1/6/2021", "1/7/2021"],
        "chart": True,
        "circle_colour": "No",
    },
    "dashboard": get_input_doc(root_path / "input/dashboards_master.xlsx"),
    "excel_save_path": str(root_path / "output/{}.xlsx"),
}

# cdg_data_dict = {
#     "docx_save_path": str(cdg_root_path / "output/{}.docx"),
#     "master": Master(
#         cdg_get_master_data(), cdg_get_project_information(), data_type="cdg"
#     ),
#     "op_args": {
#         "quarter": ["Q4 20/21"],
#         # "quarter": ["standard"],
#         "group": ["SCS", "CFPD", "GF"],
#         # "chart": True,
#         "data_type": "cdg",
#         "type": "benefits",
#         "blue_line": "CDG",
#         "dates": ["1/10/2020", "1/5/2022"],
#         "fig_size": "half_horizontal",
#     },
#     "dashboard": str(cdg_root_path / "input/dashboard_master.xlsx"),
#     "narrative_dashboard": str(cdg_root_path / "input/narrative_dashboard_master.xlsx"),
#     "excel_save_path": str(cdg_root_path / "output/{}.xlsx"),
#     "word_save_path": str(cdg_root_path / "output/{}.docx")
# }

data = ipdc_data_dict

## DANDELION
# dl_data = DandelionData(data["master"], **data["op_args"])
# d_lion = make_a_dandelion_auto(dl_data, **data["op_args"])
# put_matplotlib_fig_into_word(hoz_doc, d_lion, size=7)
# hoz_doc.save(data["docx_save_path"].format("dandelion_graph_benefits"))


# MILESTONES
ms = MilestoneData(data["master"], **data["op_args"])
ms.filter_chart_info(**data["op_args"])
wb = put_milestones_into_wb(ms)
wb.save(data["excel_save_path"].format("milestones"))
# chart_kwargs = {**{"blue_line": "today", "Chart": True}, **ms.kwargs}
# g = milestone_chart(ms, data["master"], **data["op_args"])
# put_matplotlib_fig_into_word(hoz_doc, g, size=7, transparent=False)
# hoz_doc.save(data["word_save_path"].format("milestone_graph"))

# b = BenefitsData(m, baseline=["all"])
# total_costs_benefits_bar_chart(c, b, chart=True)

# # QUERY
# # wb = data_query_into_wb(m, keys=["Senior Responsible Owner (SRO)"], quarter=["Q3 19/20"], group=DFT_GROUP)
# # wb.save(root_path / "output/query_test.xlsx")

# ## STACKPLOT
# sp_data = get_sp_data(m, quarter=["standard"], group=["RIG"])
# cost_stackplot_graph(sp_data, m, group=DFT_GROUP)

# COSTS
# c = CostData(data["master"], **data["op_args"])
# c.get_cost_profile()
# cost_profile_graph(c, data["master"], chart=True, group=c.start_group)

# SUMMARIES
# run_p_reports(data["master"], **data["op_args"])

## VFM
# c = VfMData(m, group=DFm = Master(*data["data"], **data["op_args"] )T_GROUP, quarter=["standard"])  # c is class
# wb = vfm_into_excel(c)

## RISKS
# c = RiskData(m, **op_args)
# wb = risks_into_excel(c)

# SPEED DIALS
# dca_data = DcaData(data["master"], **data["op_args"])
# dca_data.get_changes()
# build_speedials(dca_data, hoz_doc)
# hoz_doc.save(data["docx_save_path"].format("speedial_graphs"))
# doc = dca_changes_into_word(dca_data, doc)
# doc.save(data["docx_save_path"].format("speedial_dca_changes"))


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
# wb = ipdc_dashboard(data["master"], data["dashboard"], data["op_args"])
# wb.save(data["excel_save_path"].format("q4_2021_dashboard_final"))
# wb = cdg_narrative_dashboard(data["master"], data["narrative_dashboard"])
# wb.save(data["excel_save_path"].format("q4_2021_narrative_dashboard"))
