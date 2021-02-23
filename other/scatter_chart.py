import numpy as np

from analysis_engine.data import open_pickle_file, root_path, DcaData, CostData, MilestoneData, \
    cost_v_schedule_chart_into_wb, RiskData, DandelionData, VfMData, put_cost_totals_into_wb, \
    put_matplotlib_fig_into_word, open_word_doc, FIGURE_STYLE, get_cost_stackplot_data, put_stackplot_data_into_wb, \
    make_a_dandelion_auto

m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
# dandelion = DandelionData(m, quarter="standard")
d_data = DandelionData(m)
make_a_dandelion_auto(d_data)

# doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# put_dandelion_matplotlib_fig_into_word(doc, dlion, size=7.5)
# doc.save(root_path / "output/dlion_graph.docx")

# def moving_average(x, w):
#     return np.convolve(x, np.ones(w), 'valid') / w

