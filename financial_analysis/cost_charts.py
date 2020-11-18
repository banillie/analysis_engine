"""
New production code for compiling total cost and benefits bar charts.
"""

from data_mgmt.data import cost_profile_baseline_graph, \
    root_path

# master = Master(get_master_data(), get_project_information())
# master.check_baselines()
# costs = CostData(master)
# benefits = BenefitsData(master)
wd_path = root_path / "input/summary_temp_landscape.docx"

# LIST_OF_GROUPS = [master.current_projects,
#                   Projects.he,
#                   Projects.rail,
#                   Projects.rail_franchising,
#                   Projects.hs2,
#                   Projects.hsmrpg,
#                   Projects.sarh2,
#                   Projects.all_not_hs2,
#                   Projects.fbc_stage,
#                   Projects.obc_stage,
#                   Projects.sobc_stage]
# LIST_OF_TITLES = ['ALL',
#                   'HE',
#                   'RAIL INFRASTRUCTURE',
#                   'RAIL FRANCHISING',
#                   'HS2',
#                   'HSMRPG',
#                   'AMIS (SARH2)',
#                   'ALL, NOT HS2,',
#                   'FBC Projects',
#                   'OBC Projects',
#                   'SOBC Projects']


def baseline_profile(costs, project, *args):
    #costs.get_cost_profile(project, 'ipdc_costs')
    if args == ():
        cost_profile_baseline_graph(costs)
    else:
        cost_profile_baseline_graph(costs, args[0])


# def compile_all_profiles():
#     report_doc = open_word_doc(wd_path)
#     for i, p in enumerate(LIST_OF_GROUPS):
#         costs.get_cost_profile(p, 'ipdc_costs')
#         graph = cost_profile_graph(costs, LIST_OF_TITLES[i])
#         put_matplotlib_fig_into_word(report_doc, graph)
#         report_doc.save(root_path / "output/different_cost_profiles.docx")
#
#
# def compile_all_totals():
#     report_doc = open_word_doc(wd_path)
#     for i, p in enumerate(LIST_OF_GROUPS):
#         costs.get_cost_totals(p, 'ipdc_costs')
#         benefits.get_ben_totals(p, 'ipdc_benefits')
#         graph = total_costs_benefits_bar_chart(costs, benefits, LIST_OF_TITLES[i])
#         put_matplotlib_fig_into_word(report_doc, graph)
#         report_doc.save(root_path / "output/different_total_cost_profiles.docx")
#
