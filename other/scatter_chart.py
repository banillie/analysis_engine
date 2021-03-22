import math
import numpy as np
from typing import List
from matplotlib import pyplot as plt

from analysis_engine.data import (
    open_pickle_file,
    root_path,
    # DcaData,
    # CostData,
    MilestoneData,
    put_milestones_into_wb, data_query_into_wb, milestone_chart,
    # cost_v_schedule_chart_into_wb,
    # RiskData,
    # DandelionData,
    # VfMData,
    # put_cost_totals_into_wb,
    put_matplotlib_fig_into_word,
    open_word_doc,
    # FIGURE_STYLE,
    # get_cost_stackplot_data,
    # put_stackplot_data_into_wb,
    # make_a_dandelion_auto,
    # BenefitsData,
    # total_costs_benefits_bar_chart,
    # gauge,
    # calculate_arg_combinations
)

m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
# arg_list = ["quarter", "baseline", "group", "stage", "type", "remove", "dates", "koi"]
# calculate_arg_combinations(arg_list)
# stage = ["pre-SOBC", "SOBC", "OBC", "FBC"]
# group = ["HSMRPG", "AMIS", "Rail", "RPE"]
# d_data = DandelionData(m, quarter=["Q3 20/21"], group=group, remove=["Rail Franchising", "Crossrail"])
# d_lion = make_a_dandelion_auto(d_data, title="Standard dandelion", chart=True)
# doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# put_matplotlib_fig_into_word(doc, d_lion, size=7.5)
# doc.save(root_path / "output/dlion_graph.docx")

## MILESTONES
ms = MilestoneData(m, quarter=[str(m.current_quarter)], group=["GMPP"])
ms.filter_chart_info(dates=["1/1/2021", "1/1/2030"])
wb = put_milestones_into_wb(ms)
wb.save(root_path / "output/gmpp_milestones_data.xlsx")
doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
for p in m.dft_groups["Q3 20/21"]["GMPP"]:
    try:
        ms = MilestoneData(m, quarter=[str(m.current_quarter)], group=[p])
        ms.filter_chart_info(dates=["1/1/2021", "1/1/2030"])
        graph = milestone_chart(ms, blue_line="Today", chart=True)
        put_matplotlib_fig_into_word(doc, graph, size=8, transparent=False)
    except ValueError:
        pass

doc.save(root_path / "output/gmpp_milestones_charts.docx".format(p))

# b = BenefitsData(m, baseline=["all"])
# total_costs_benefits_bar_chart(c, b, chart=True)

## QUERY
wb = data_query_into_wb(m, keys=["FBC - IPDC Approval"], quarter=["all"])
wb.save(root_path / "output/query_test.xlsx")

