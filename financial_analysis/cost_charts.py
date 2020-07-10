from data_mgmt.data import list_of_masters_all, root_path, Masters, CostData, \
    BenefitsData, hs2_programme, rail_franchising, rpe, lower_thames_crossing
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import datetime
from datetime import timedelta
from textwrap import wrap

test_masters = list_of_masters_all[1:]
all_projects = list_of_masters_all[1].projects
all_projects.remove(hs2_programme)
all_projects.remove(rail_franchising)

mst = Masters(test_masters, all_projects)
mst.baseline_data('Re-baseline IPDC cost')
costs = CostData(mst)
costs.cost_totals([lower_thames_crossing])
#mst.baseline_data('Re-baseline IPDC cost')
bens = BenefitsData(mst)
bens.ben_totals([lower_thames_crossing])


def cost_charts(costs_obj, bens_obj, graph_title):

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
    # benefits change
    labels = ['Baseline', 'Last Quarter', 'Latest']
    width = 0.5
    ax2.bar(labels, np.array(bens_obj.achieved), width, label='Delivered')
    ax2.bar(labels, np.array(bens_obj.profile), width, bottom=np.array(bens_obj.achieved), label='Profiled')
    ax2.bar(labels, np.array(bens_obj.unprofile), width,
            bottom=np.array(bens_obj.achieved) + np.array(bens_obj.profile), label='Unprofiled')
    ax2.legend(prop={'size': 6})
    ax2.set_ylabel('Benefits (£m)')
    ylab2 = ax2.yaxis.get_label()
    ylab2.set_style('italic')
    ylab2.set_size(8)
    ax2.tick_params(axis='x', which='major', labelsize=6)
    ax2.tick_params(axis='y', which='major', labelsize=6)
    ax2.set_title('Fig 3 - ben total change over time', loc='left', fontsize=8, fontweight='bold')

    ax2.set_ylim(0, y_max)


    # benefits break down
    labels = ['Cashable', 'Non-Cashable', 'Economic', 'Disbenefit']
    width = 0.5
    ax4.bar(labels, np.array(bens_obj.cat_achieved), width, label='Delivered')
    ax4.bar(labels, np.array(bens_obj.cat_profile), width, bottom=np.array(bens_obj.cat_achieved), label='Profiled')
    ax4.bar(labels, np.array(bens_obj.cat_unprofile), width,
            bottom=np.array(bens_obj.cat_achieved) + np.array(bens_obj.cat_profile), label='Unprofiled')
    ax4.legend(prop={'size': 6})
    ax4.set_ylabel('Benefits (£m)')
    ylab4 = ax4.yaxis.get_label()
    ylab4.set_style('italic')
    ylab4.set_size(8)
    ax4.tick_params(axis='x', which='major', labelsize=6)
    ax4.tick_params(axis='y', which='major', labelsize=6)
    ax4.set_title('Fig 4 - benefits profile type', loc='left', fontsize=8, fontweight='bold')

    y_min = min(-100000)
    ax4.set_ylim(y_min, y_max)

    # size of chart and fit
    fig.canvas.draw()
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

    fig.savefig(root_path / 'output/{}.png'.format(graph_title), bbox_inches='tight')
    #plt.close()  # automatically closes figure so don't need to do manually.

    #doc.add_picture('cost_bens_overview.png', width=Inches(8))  # to place nicely in doc
    #os.remove('cost_bens_overview.png')


cost_charts(costs, bens, 'test')

