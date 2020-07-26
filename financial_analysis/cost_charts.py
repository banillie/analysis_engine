from data_mgmt.data import list_of_masters_all, root_path, Masters, CostData, \
    BenefitsData
from data_mgmt import Projects
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import datetime
from datetime import timedelta
from textwrap import wrap

test_masters = list_of_masters_all
all_projects = list_of_masters_all[0].projects
#all_projects.remove(ProjectGroupName.hs2_programme)
# all_projects.remove(Projects.hs2_1)
# all_projects.remove(Projects.hs2_2a)
# all_projects.remove(Projects.hs2_2b)
#all_projects.remove(ProjectGroupName.tru)

year_list = ['20-21',
             '21-22',
             '22-23',
             '23-24',
             '24-25',
             '25-26',
             '26-27',
             '27-28',
             '28-29']

mst = Masters(test_masters, all_projects)
mst.baseline_data('Re-baseline IPDC cost')
costs = CostData(mst)
mst.baseline_data('Re-baseline IPDC benefits')
bens = BenefitsData(mst)


def cost_charts(costs_obj, bens_obj, graph_title):

    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2) #four sub plotsprint

    #fig.suptitle(abbreviations[project_name] + ' costs and benefits analysis', fontweight='bold') # title

    #Cost changes over time bar chart
    labels = ['Baseline', 'Last Quarter', 'Current']
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

    #scaling y axis. y axis value setting so it takes either highest ben or cost figure
    max_value = max(costs_obj.total + bens_obj.total)
    y_max = max_value + max_value / 5
    ax1.set_ylim(0, y_max)

    #cost types/categories bar chart
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
    ax3.set_ylim(0, y_max) #scale y axis max

    # benefits change over time bar chart
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


    # benefits type/category bar chart
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

    y_min = min(bens_obj.disbenefit)
    ax4.set_ylim(y_min, y_max)

    # size of chart and fit
    fig.canvas.draw()
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

    fig.savefig(root_path / 'output/{}.png'.format(graph_title), bbox_inches='tight')
    #plt.close()  # automatically closes figure so don't need to do manually.

    #doc.add_picture('cost_bens_overview.png', width=Inches(8))  # to place nicely in doc
    #os.remove('cost_bens_overview.png')

def cost_profile_chart(costs_obj, graph_title):

    fig, ax1 = plt.subplots() #four sub plotsprint

    '''cost profile charts'''
    year = year_list
    #fig.suptitle(abbreviations[project_name] + ' financial analysis', fontweight='bold') # title

    #plot cost change profile chart
    ax1.plot(year, np.array(costs_obj.baseline_profile), label='Baseline', linewidth=3.0, marker="o")
    #ax1.plot(year, np.array(costs_obj.last_profile), label='Last quarter', linewidth=3.0, marker="o")
    ax1.plot(year, np.array(costs_obj.current_profile), label='Latest', linewidth=3.0, marker="o")

    #cost profile change chart styling
    ax1.tick_params(axis='x', which='major', labelsize=6, rotation=45)
    ax1.set_ylabel('Cost (£m)')
    ylab1 = ax1.yaxis.get_label()
    ylab1.set_style('italic')
    ylab1.set_size(8)
    ax1.grid(color='grey', linestyle='-', linewidth=0.2)
    ax1.legend(prop={'size': 6})
    ax1.set_title('Fig 1 - cost profile changes', loc='left', fontsize=8, fontweight='bold')

    # scaling y axis
    # y axis value setting so it takes highest cost profile yeah
    all = costs_obj.current_profile + costs_obj.last_profile + costs_obj.baseline_profile
    y_max = max(all) + max(all) * 1 / 5
    ax1.set_ylim(0, y_max)

    # # plot rdel/cdel chart data
    # ax2.plot(year, latest_profile_cdel, label='CDEL', linewidth=3.0, marker="o")
    # ax2.plot(year, latest_profile_rdel, label='RDEL', linewidth=3.0, marker="o")
    #
    # #rdel/cdel profile chart styling
    # ax2.tick_params(axis='x', which='major', labelsize=6, rotation=45)
    # ax2.set_xlabel('Financial Years')
    # ax2.set_ylabel('Cost (£m)')
    # xlab2 = ax2.xaxis.get_label()
    # ylab2 = ax2.yaxis.get_label()
    # xlab2.set_style('italic')
    # xlab2.set_size(8)
    # ylab2.set_style('italic')
    # ylab2.set_size(8)
    # ax2.grid(color='grey', linestyle='-', linewidth=0.2)
    # ax2.legend(prop={'size': 6})
    #
    # ax2.set_ylim(0, y_max)
    #
    # ax2.set_title('Fig 2 - cost profile spend type', loc='left', fontsize=8, fontweight='bold')

    # size of chart and fit
    fig.canvas.draw()
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

    fig.savefig(root_path / 'output/{}.png'.format(graph_title), bbox_inches='tight')
    # plt.close()  # automatically closes figure so don't need to do manually.

    # doc.add_picture('cost_profile.png', width=Inches(8))  # to place nicely in doc
    # os.remove('cost_profile.png')

def report_cost_profile_chart(costs_obj, graph_title, style):

    plt.style.use(style)
    fig, ax1 = plt.subplots()
    fig.set_figheight(4)
    fig.set_figwidth(8)

    '''cost profile charts'''
    year = year_list

    #fig.suptitle(abbreviations[project_name] + ' financial analysis', fontweight='bold') # title

    #plot cost change profile chart
    ax1.plot(year, np.array(costs_obj.baseline_profile), label='Baseline', linewidth=3.0, marker="o")
    ax1.plot(year, np.array(costs_obj.last_profile), label='Last', linewidth=3.0, marker="o")
    ax1.plot(year, np.array(costs_obj.current_profile), label='Current', linewidth=3.0, marker="o")

    #cost profile change chart styling
    # ax1.tick_params(axis='x', which='major', labelsize=6, rotation=45)
    ax1.set_ylabel('Cost (£m)')
    ylab1 = ax1.yaxis.get_label()
    ylab1.set_style('italic')
    ax1.set_xlabel('Financial Year')
    xlab1 = ax1.xaxis.get_label()
    xlab1.set_style('italic')

    #ylab1.set_size(8)
    # ax1.grid(color='grey', linestyle='-', linewidth=0.2)
    #ax1.legend(prop={'size': 6})
    ax1.legend()
    # ax1.set_title('Fig 1 - cost profile changes', loc='left', fontsize=8, fontweight='bold')

    # scaling y axis
    # y axis value setting so it takes highest cost profile yeah
    # all = costs_obj.current_profile + costs_obj.last_profile + costs_obj.baseline_profile
    # y_max = max(all) + max(all) * 1 / 5
    # ax1.set_ylim(0, y_max)

    # size of chart and fit
    fig.canvas.draw()
    # fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

    fig.savefig(root_path / 'output/{}.png'.format(graph_title), bbox_inches='tight')
    # plt.close()  # automatically closes figure so don't need to do manually.

    # doc.add_picture('cost_profile.png', width=Inches(8))  # to place nicely in doc
    # os.remove('cost_profile.png')

# cost_charts(costs, bens, 'test')

#cost_profile_chart(costs, 'testies')

#styles = plt.style.available
# styles = ['seaborn']
# for style in styles:
#     report_cost_profile_chart(costs, style, style)