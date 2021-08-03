import json
import math
from dateutil import parser

import numpy as np

# from typing import List
from matplotlib import pyplot as plt
from openpyxl import load_workbook

# matplotlib==3.3.0
from analysis_engine.cdg_data import (
    # cdg_root_path,
    # cdg_get_master_data,
    # cdg_get_project_information,
    cdg_dashboard,
    cdg_run_p_reports,
    cdg_narrative_dashboard,
)
from analysis_engine.data import (
    cdg_root_path,
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
    JsonMaster,
    get_gmpp_data,
    get_risk_data,
    print_risk_data,
    get_map,
    get_project_info_data,
    data_check_print_out,
    get_ipdc_date,
    doughut,
    dca_changes_into_excel,
    get_cost_forecast_keys,
    cost_profile_graph_new,
    cost_profile_into_wb_new,
    put_n02_into_master,
    data_query_into_wb_by_key,
    amend_project_information,
    compile_p_report,
    compile_p_report_new,
)

## GENERATE CLI OPTIONS
# arg_list = ["quarters", "group", "stage", "remove", "type"]
# calculate_arg_combinations(arg_list)

from datamaps.api import project_data_from_master_month, project_data_from_master

# from analysis_engine.top35_data import (
#     top35_root_path,
#     top35_run_p_reports, CentralSupportData,
# )

# INITIATE
# master = JsonMaster(
#     get_master_data(
#         str(cdg_root_path) + "/core_data/cdg_config.ini",
#         str(cdg_root_path) + "/core_data/",
#         project_data_from_master
#     ),
#     get_project_information(
#         str(cdg_root_path) + "/core_data/cdg_config.ini",
#         str(cdg_root_path) + "/core_data/"
#     ),
#     data_type="cdg"
# )
# master_json_path = str("{0}/core_data/json/master".format(cdg_root_path))
# JsonData(master, master_json_path)


## GROUPS
# DFT_GROUP = ["HSRG", "RSS", "RIG", "AMIS", "RPE"]
# STAGE_GROUPS = ["pipeline", "pre-SOBC", "SOBC", "OBC", "FBC"]
hoz_doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# doc = open_word_doc(root_path / "input/summary_temp.docx")

# top35_data_dict = {
#     "docx_save_path": str(top35_root_path / "output/{}.docx"),
#     "master": Master(open_json_file(str(top35_root_path / "core_data/json/master.json"))),
#     "op_args": {
#         # "quarter": ["Month(June), 2021"],
#         "quarter": ["standard"],
#         # "group": ["HSRG", "RSS", "RIG", "RPE"],
#         "group": ["LIC"],
#         # "chart": False,
#         "data_type": "top35",
#         "circle_colour": "No",
#         # "dates": ["1/6/2021", "1/7/2021"],
#         "key": [
#             "PROJECT DEL TO CURRENT TIMINGS ?",
#             "GMPP ID: IS THIS PROJECT ON GMPP",
#             "PROJECT ON BUDGET?",
#             "WLC TOTAL",
#             "WLC NON GOV",
#         ],
#         # "key": ["Start of Trial Running"],
#     },
#     "excel_save_path": str(top35_root_path / "output/{}.xlsx"),
#     "word_save_path": str(top35_root_path / "output/{}.docx")
# }

ipdc_data_dict = {
    "docx_save_path": str(root_path / "output/{}.docx"),
    "master": Master(open_json_file(str(root_path / "core_data/json/master.json"))),
    "op_args": {
        # "quarter": ["standard"],
        "quarter": ["Q1 21/22"],
        # "baseline": ["current"],
        # "group": ["HSRG", "RSS", "RIG", "AMIS", "RPE"],
        "group": ["South West Route Capacity"],
        # "stage": ["pre-SOBC", "SOBC", "OBC", "FBC"],
        # "remove": ["HS2 Ph 2b", "HS2 Ph 2a", "NPR"],
        # "dates": ["1/6/2021", "1/7/2021"],
        # "type": "short",
        # "chart": True,
        # "circle_colour": "No",
        # "key": ["VfM Category single entry"],
        # "conf_type": "sro",
        # "rag_number": "3",
        # "order_by": "schedule",
        # "angles": [240, 290, 22, 120],
        # "weighting": "count",
        # "show": "No",
    },
    "dashboard": get_input_doc(root_path / "input/dashboards_master.xlsx"),
    "excel_save_path": str(root_path / "output/{}.xlsx"),
    "word_save_path": str(root_path / "output/{}.docx"),
}

# cdg_data_dict = {
#     "docx_save_path": str(cdg_root_path / "output/{}.docx"),
#     "master": Master(open_json_file(str(cdg_root_path / "core_data/json/master.json"))),
#     "op_args": {
#         "quarter": ["Q1 21/22"],
#         # "quarter": ["standard"],
#         "group": ["SCS", "CFPD", "GF"],
#         # "group": ["SCS", "GF"],
#         "chart": True,
#         "data_type": "cdg",
#         "type": "income",
#         "blue_line": "CDG",
#         "dates": ["1/3/2021", "1/6/2022"],
#         "fig_size": "half_horizontal",
#         "rag_number": "5",
#         # "order_by": "cost",
#         "angles": [300, 360, 60],
#         "none_handle": "none",
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
# hoz_doc.save(data["word_save_path"].format("dandelion_income"))


# MILESTONES
# ms = MilestoneData(data["master"], **data["op_args"])
# ms.filter_chart_info(**data["op_args"])
# wb = put_milestones_into_wb(ms)
# wb.save(data["excel_save_path"].format("milestones"))
# g = milestone_chart(ms, data["master"], **data["op_args"])
# put_matplotlib_fig_into_word(hoz_doc, g, size=7, transparent=False)
# hoz_doc.save(data["word_save_path"].format("milestone_graph"))

# b = BenefitsData(m, baseline=["all"])
# total_costs_benefits_bar_chart(c, b, chart=True)

# QUERY
# wb = data_query_into_wb_by_key(data["master"], **data["op_args"])
# wb.save(data["excel_save_path"].format("query_vfm_single_cat"))

# STACKPLOT
# sp_data = get_sp_data(m, quarter=["standard"], group=["RIG"])
# cost_stackplot_graph(sp_data, m, group=DFT_GROUP)

# COSTS

# for x in data["master"].current_projects:
#     data["op_args"]["group"] = [x]
# c = CostData(data["master"], **data["op_args"])
# c.get_baseline_cost_profile()
# c.get_forecast_cost_profile()
# g = cost_profile_graph_new(c, data["master"], **data["op_args"])
# put_matplotlib_fig_into_word(hoz_doc, g, size=7, transparent=False)
# hoz_doc.save(data["word_save_path"].format("bl_portfolio_no_npr_hs22a2b"))
# wb = cost_profile_into_wb_new(c)
# wb.save(data["excel_save_path"].format("costs_q1_2021"))
# wb.save(data["excel_save_path"].format(str(x) + " costs_q1_2021"))

# SUMMARIES
# run_p_reports(data["master"], **data["op_args"])

## VFM
# c = VfMData(m, group=DFm = Master(*data["data"], **data["op_args"] )T_GROUP, quarter=["standard"])  # c is class
# wb = vfm_into_excel(c)

## RISKS
c = RiskData(data["master"], **data["op_args"])
wb = risks_into_excel(c)
wb.save(data["excel_save_path"].format("risks"))

# SPEED DIALS
# dca_data = DcaData(data["master"], **data["op_args"])
# dca_data.get_changes()
# build_speedials(dca_data, hoz_doc)
# hoz_doc.save(data["docx_save_path"].format("speedial_graphs"))
# doc = dca_changes_into_word(dca_data, doc)
# doc.save(data["docx_save_path"].format("speedial_dca_changes"))

# DCAS
# wb = dca_changes_into_excel(dca_data)
# wb.save(data["excel_save_path"].format("dcas_testing"))

# DOUGHUTS
# dough = doughut(dca_data, **data["op_args"])
# put_matplotlib_fig_into_word(hoz_doc, dough, size=7.5)
# hoz_doc.save(data["docx_save_path"].format("sro_count_doughut"))

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
# wb.save(data["excel_save_path"].format("ipdc_dashboard_testing"))
# wb = cdg_narrative_dashboard(data["master"], data["narrative_dashboard"])
# wb.save(data["excel_save_path"].format("q1_2021_narrative_dashboard"))
# wb = cdg_dashboard(data["master"], data["dashboard"])
# wb.save(data["excel_save_path"].format("q1_2021_dashboard"))

##CENTRAL SUPPORT
# cs = CentralSupportData(data["master"], **data["op_args"])
# wb = put_milestones_into_wb(cs)
# wb.save(data["excel_save_path"].format("central_support_testing"))
# g = milestone_chart(cs, data["master"], **data["op_args"])
# put_matplotlib_fig_into_word(hoz_doc, g, size=7, transparent=False)
# hoz_doc.save(data["word_save_path"].format("central_support_graph"))

# GMPP data
# get_gmpp_data()
# gmpp_d = get_project_info_data("/home/will/Downloads/GMPP_DATA_DFT_FORMAT_Q1_FINAL.xlsx")
# ipdc_d = get_project_info_data(root_path / "core_data/master_1_2021.xlsx")
# data_check_print_out(gmpp_d, ipdc_d)
# change_gmpp_keys_order(ipdc_d)
# put_n02_into_master()

# Risk data
# risks = get_risk_data()
# print_risk_data(risks)

# work for cost data. to delete.
# from analysis_engine.data import get_cost_baseline_keys
# t = get_cost_forecast_keys()


# amend_project_information()
