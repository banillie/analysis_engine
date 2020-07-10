from data_mgmt.data import list_of_masters_all, root_path, Masters, CostData, \
    hs2_programme, rail_franchising
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import datetime
from datetime import timedelta
from textwrap import wrap

test_masters = list_of_masters_all[1:]
project_names = list_of_masters_all[1].projects
project_names.remove(hs2_programme)
project_names.remove(rail_franchising)

mst = Masters(test_masters, project_names)
mst.get_baseline_data('Re-baseline IPDC cost')
costs = CostData(mst)

def cost_charts(costs_obj, graph_title):

    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2) #four sub plotsprint

    #fig.suptitle(abbreviations[project_name] + ' costs and benefits analysis', fontweight='bold') # title

    #Spent, Profiled and Unprofile chart
    labels = ['Baseline', 'Last Quarter', 'Current']
        #['Current', 'Last quarter', 'Baseline']
    width = 0.5
    #matplotlib requires lists to numpy arrays
    ax1.bar(labels, np.array(costs_obj.spent), width, label='Spent')
    ax1.bar(labels, np.array(costs_obj.profile), width, bottom=np.array(costs_obj.spent), label='Profiled')
    ax1.bar(labels, np.array(costs_obj.unprofile), width,
            bottom=np.array(costs_obj.spent) + np.array(costs_obj.profile), label='Unprofiled')
    ax1.legend(prop={'size': 6})
    ax1.set_ylabel('Cost (£m)')
    ylab1 = ax1.yaxis.get_label()
    ylab1.set_style('italic')
    ylab1.set_size(8)
    ax1.tick_params(axis='x', which='major', labelsize=6)
    ax1.tick_params(axis='y', which='major', labelsize=6)
    ax1.set_title('Fig 1 - total costs change over time', loc='left', fontsize=8, fontweight='bold')

    #scaling y axis
    #y axis value setting so it takes either highest ben or cost figure
    cost_max = max(costs_obj.profile)*3 #+ max(costs_obj.profile)/2
    #ben_max = max(total_ben) + max(total_ben)/5
    y_max = cost_max
    #y_max = max([cost_max, ben_max])
    ax1.set_ylim(0, y_max)

    #rdel/cdel bar chart

    labels = ['RDEL', 'CDEL']
    width = 0.5
    ax3.bar(labels, np.array(costs_obj.cat_spent), width, label='Spent')
    ax3.bar(labels, np.array(costs_obj.cat_profile), width, bottom=np.array(costs_obj.cat_spent), label='Profiled')
    ax3.bar(labels, np.array(costs_obj.cat_unprofiled), width,
            bottom=np.array(costs_obj.cat_spent) + np.array(costs_obj.cat_profile), label='Unprofiled')
    ax3.legend(prop={'size': 6})
    ax3.set_ylabel('Costs (£m)')
    ylab3 = ax3.yaxis.get_label()
    ylab3.set_style('italic')
    ylab3.set_size(8)
    ax3.tick_params(axis='x', which='major', labelsize=6)
    ax3.tick_params(axis='y', which='major', labelsize=6)
    ax3.set_title('Fig 2 - current costs category break down', loc='left', fontsize=8, fontweight='bold')

    #y_max = max(total_fin) + max(total_fin) * 1 / 5
    ax3.set_ylim(0, y_max) #scale y axis max
    #
    # # benefits change
    # labels = ['Baseline', 'Last Quarter', 'Latest']
    # width = 0.5
    # ax2.bar(labels, delivered_ben, width, label='Delivered')
    # ax2.bar(labels, profiled_ben, width, bottom=delivered_ben, label='Profiled')
    # ax2.bar(labels, unprofiled_ben, width, bottom=delivered_ben + profiled_ben, label='Unprofiled')
    # ax2.legend(prop={'size': 6})
    # ax2.set_ylabel('Benefits (£m)')
    # ylab2 = ax2.yaxis.get_label()
    # ylab2.set_style('italic')
    # ylab2.set_size(8)
    # ax2.tick_params(axis='x', which='major', labelsize=6)
    # ax2.tick_params(axis='y', which='major', labelsize=6)
    # ax2.set_title('Fig 3 - ben total change over time', loc='left', fontsize=8, fontweight='bold')
    #
    # ax2.set_ylim(0, y_max)
    #
    # # benefits break down
    # labels = ['Cashable', 'Non-Cashable', 'Economic', 'Disbenefit']
    # width = 0.5
    # ax4.bar(labels, type_delivered_ben, width, label='Delivered')
    # ax4.bar(labels, type_profiled_ben, width, bottom=type_delivered_ben, label='Profiled')
    # ax4.bar(labels, type_unprofiled_ben, width, bottom=type_delivered_ben + type_profiled_ben, label='Unprofiled')
    # ax4.legend(prop={'size': 6})
    # ax4.set_ylabel('Benefits (£m)')
    # ylab4 = ax4.yaxis.get_label()
    # ylab4.set_style('italic')
    # ylab4.set_size(8)
    # ax4.tick_params(axis='x', which='major', labelsize=6)
    # ax4.tick_params(axis='y', which='major', labelsize=6)
    # ax4.set_title('Fig 4 - benefits profile type', loc='left', fontsize=8, fontweight='bold')
    #
    # y_min = min(type_disbenefit_ben)
    # ax4.set_ylim(y_min, y_max)
    #
    # size of chart and fit
    fig.canvas.draw()
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

    fig.savefig(root_path / 'output/{}.png'.format(graph_title), bbox_inches='tight')
    #plt.close()  # automatically closes figure so don't need to do manually.

    #doc.add_picture('cost_bens_overview.png', width=Inches(8))  # to place nicely in doc
    #os.remove('cost_bens_overview.png')


cost_charts(costs, 'test')


# def cost_profile_graph(mst, c):
#
#     fig, (ax1, ax2) = plt.subplots(2) #four sub plotsprint
#
#     '''cost profile charts'''
#     year = year_list[:-1]
#     baseline_profile_total = profile_data_total[2]
#     last_profile_total = profile_data_total[1]
#     latest_profile_total = profile_data_total[0]
#
#     latest_profile_rdel = profile_data_rdel[0]
#     latest_profile_cdel = profile_data_cdel[0]
#
#     fig.suptitle(abbreviations[project_name] + ' financial analysis', fontweight='bold') # title
#
#     #plot cost change profile chart
#     ax1.plot(year, baseline_profile_total, label='Baseline', linewidth=3.0, marker="o")
#     ax1.plot(year, last_profile_total, label='Last quarter', linewidth=3.0, marker="o")
#     ax1.plot(year, latest_profile_total, label='Latest', linewidth=3.0, marker="o")
#
#     #cost profile change chart styling
#     ax1.tick_params(axis='x', which='major', labelsize=6, rotation=45)
#     ax1.set_ylabel('Cost (£m)')
#     ylab1 = ax1.yaxis.get_label()
#     ylab1.set_style('italic')
#     ylab1.set_size(8)
#     ax1.grid(color='grey', linestyle='-', linewidth=0.2)
#     ax1.legend(prop={'size': 6})
#     ax1.set_title('Fig 1 - cost profile changes', loc='left', fontsize=8, fontweight='bold')
#
#     # scaling y axis
#     # y axis value setting so it takes highest cost profile yeah
#     all = profile_data_total[0] + profile_data_total[1] + profile_data_total[2]
#     y_max = max(all) + max(all) * 1 / 5
#     ax1.set_ylim(0, y_max)
#
#     # plot rdel/cdel chart data
#     ax2.plot(year, latest_profile_cdel, label='CDEL', linewidth=3.0, marker="o")
#     ax2.plot(year, latest_profile_rdel, label='RDEL', linewidth=3.0, marker="o")
#
#     #rdel/cdel profile chart styling
#     ax2.tick_params(axis='x', which='major', labelsize=6, rotation=45)
#     ax2.set_xlabel('Financial Years')
#     ax2.set_ylabel('Cost (£m)')
#     xlab2 = ax2.xaxis.get_label()
#     ylab2 = ax2.yaxis.get_label()
#     xlab2.set_style('italic')
#     xlab2.set_size(8)
#     ylab2.set_style('italic')
#     ylab2.set_size(8)
#     ax2.grid(color='grey', linestyle='-', linewidth=0.2)
#     ax2.legend(prop={'size': 6})
#
#     ax2.set_ylim(0, y_max)
#
#     ax2.set_title('Fig 2 - cost profile spend type', loc='left', fontsize=8, fontweight='bold')
#
#     # size of chart and fit
#     fig.canvas.draw()
#     fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title
#
#     fig.savefig('cost_profile.png')
#     plt.close()  # automatically closes figure so don't need to do manually.
#
#     doc.add_picture('cost_profile.png', width=Inches(8))  # to place nicely in doc
#     os.remove('cost_profile.png')
#
#     return doc
