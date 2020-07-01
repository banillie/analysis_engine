from analysis.data import list_of_masters_all, abbreviations, ipdc_date, root_path, hsmrpg, iep, hs2_programme
from data_mgmt.data import MilestoneData, MilestoneChartData, Baselines, MilestoneCharts
import datetime
from datetime import timedelta
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
from textwrap import wrap




# def milestone_swimlane_charts(latest_milestone_names,
#                               latest_milestone_dates,
#                               last_milestone_dates,
#                               baseline_milestone_dates,
#                               graph_title):
#
#     #build scatter chart
#     fig, ax1 = plt.subplots()
#     fig.suptitle(graph_title, fontweight='bold')  # title
#     # set fig size
#     fig.set_figheight(4)
#     fig.set_figwidth(8)
#
#     ax1.scatter(baseline_milestone_dates, latest_milestone_names, label='Baseline')
#     ax1.scatter(last_milestone_dates, latest_milestone_names, label='Last Qrt')
#     ax1.scatter(latest_milestone_dates, latest_milestone_names, label='Latest Qrt')
#
#     # format the x ticks
#     years = mdates.YearLocator()  # every year
#     months = mdates.MonthLocator()  # every month
#     years_fmt = mdates.DateFormatter('%Y')
#     months_fmt = mdates.DateFormatter('%b')
#
#     # calculate the length of the time period covered in chart. Not perfect as baseline dates can distort.
#     try:
#         td = (latest_milestone_dates[-1] - latest_milestone_dates[0]).days
#         if td <= 365*3:
#             ax1.xaxis.set_major_locator(years)
#             ax1.xaxis.set_minor_locator(months)
#             ax1.xaxis.set_major_formatter(years_fmt)
#             ax1.xaxis.set_minor_formatter(months_fmt)
#             plt.setp(ax1.xaxis.get_minorticklabels(), rotation=45)
#             plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, weight='bold')# milestone_swimlane_charts(key_name,
# #                           current_m_data,
# #                           last_m_data,
# #                           baseline_m_data,
# #                           'All Milestones')
#             # scaling x axis
#             # x axis value to no more than three months after last latest milestone date, or three months
#             # before first latest milestone date. Hack, can be improved. Text highlights movements off chart.
#             x_max = latest_milestone_dates[-1] + timedelta(days=90)
#             x_min = latest_milestone_dates[0] - timedelta(days=90)
#             for date in baseline_milestone_dates:
#                 if date > x_max:
#                     ax1.set_xlim(x_min, x_max)
#                     plt.figtext(0.98, 0.03,
#                                 'Check full schedule to see all milestone movements',
#                                 horizontalalignment='right', fontsize=6, fontweight='bold')
#                 if date < x_min:
#                     ax1.set_xlim(x_min, x_max)
#                     plt.figtext(0.98, 0.03,
#                                 'Check full schedule to see all milestone movements',
#                                 horizontalalignment='right', fontsize=6, fontweight='bold')
#         else:
#             ax1.xaxis.set_major_locator(years)
#             ax1.xaxis.set_minor_locator(months)
#             ax1.xaxis.set_major_formatter(years_fmt)
#             plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, weight='bold')
#     except IndexError: #if milestone dates list is empty:
#         pass
#
#     ax1.legend() #insert legend
#
#     #reverse y axis so order is earliest to oldest
#     ax1 = plt.gca()
#     ax1.set_ylim(ax1.get_ylim()[::-1])
#     ax1.tick_params(axis='y', which='major', labelsize=7)
#     ax1.yaxis.grid()  # horizontal lines
#     ax1.set_axisbelow(True)
#     #ax1.get_yaxis().set_visible(False)
#
#     # for i, txt in enumerate(latest_milestone_names):
#     #     ax1.annotate(txt, (i, latest_milestone_dates[i]))
#
#     #Add line of IPDC date, but only if in the time period
#     try:
#         if latest_milestone_dates[0] <= ipdc_date <= latest_milestone_dates[-1]:
#             plt.axvline(ipdc_date)
#             plt.figtext(0.98, 0.01, 'Line represents when IPDC will discuss Q1 20_21 portfolio management report',
#                         horizontalalignment='right', fontsize=6, fontweight='bold')
#     except IndexError:
#         pass
#
#     #size of chart and fit
#     fig.canvas.draw()
#     fig.tight_layout(rect=[0, 0.03, 1, 0.95]) #for title
#
#     fig.savefig(root_path/'output/{}.png'.format(graph_title), bbox_inches='tight')
#
#     #plt.close() #automatically closes figure so don't need to do manually.

def build_charts(latest_milestone_names,
                 latest_milestone_dates,
                 last_milestone_dates,
                 baseline_milestone_dates,
                 graph_title,
                 ipdc_date):

    # add \n to y axis labels and cut down if two long
    #labels = ['\n'.join(wrap(l, 40)) for l in latest_milestone_names]
    labels = latest_milestone_names
    final_labels = []
    for l in labels:
        if len(l) > 40:
            final_labels.append(l[:35])
        else:
            final_labels.append(l)

    # Chart
    no_milestones = len(latest_milestone_names)

    if no_milestones <= 30:
        MilestoneCharts.milestone_swimlane_charts(np.array(final_labels), np.array(latest_milestone_dates),
                                  np.array(last_milestone_dates),
                                  np.array(baseline_milestone_dates), graph_title, ipdc_date)

    if 31 <= no_milestones <= 60:
        half = int(no_milestones / 2)
        MilestoneCharts.milestone_swimlane_charts(np.array(final_labels[:half]),
                                  np.array(latest_milestone_dates[:half]),
                                  np.array(last_milestone_dates[:half]), np.array(baseline_milestone_dates[:half]),
                                  graph_title, ipdc_date)
        title = graph_title + ' cont.'
        MilestoneCharts.milestone_swimlane_charts(np.array(final_labels[half:no_milestones]),
                                  np.array(latest_milestone_dates[half:no_milestones]),
                                  np.array(last_milestone_dates[half:no_milestones]),
                                  np.array(baseline_milestone_dates[half:no_milestones]), title, ipdc_date)

    if 61 <= no_milestones <= 90:
        third = int(no_milestones / 3)
        MilestoneCharts.milestone_swimlane_charts(np.array(final_labels[:third]),
                                  np.array(latest_milestone_dates[:third]),
                                  np.array(last_milestone_dates[:third]),
                                  np.array(baseline_milestone_dates[:third]),
                                  graph_title, ipdc_date)
        title = graph_title + ' cont. 1'
        MilestoneCharts.milestone_swimlane_charts(np.array(final_labels[third:third * 2]),
                                  np.array(latest_milestone_dates[third:third * 2]),
                                  np.array(last_milestone_dates[third:third * 2]),
                                  np.array(baseline_milestone_dates[third:third * 2]),
                                  title, ipdc_date)
        title = graph_title + ' cont. 2'
        MilestoneCharts(np.array(final_labels[third * 2:no_milestones]),
                                  np.array(latest_milestone_dates[third * 2:no_milestones]),
                                  np.array(last_milestone_dates[third * 2:no_milestones]),
                                  np.array(baseline_milestone_dates[third * 2:no_milestones]),
                                  title, ipdc_date)
    pass

#get data via import of list_of_masters_all
#get list of project names
p_n_list = list_of_masters_all[0].projects
p_n_list.remove(iep)
p_n_list.remove(hs2_programme)

#get baselines
b = Baselines(list_of_masters_all)
b_i = b.get_baseline_data(p_n_list) # get_baseline_data returning two things

# get general milestone data
m = MilestoneData(list_of_masters_all, baseline_index=b_i[1]) #what does baseline_index = do?
current_m = m.group_data(p_n_list, 0, abbreviations)
last_m = m.group_data(p_n_list, 1, abbreviations)
baseline_m = m.group_data(p_n_list, 2, abbreviations)

#get milestone chart data
m_chart = MilestoneChartData(current_m, last_m, baseline_m)

#filtering data options for the chart
start_date = datetime.date(2020, 3, 1)
end_date = datetime.date(2022, 6, 30)

parliament = ['Bill', 'bill', 'hybrid', 'Hybrid', 'reading',
              'royal', 'Royal', 'assent', 'Assent',
              'legislation', 'Legislation', 'Passed', 'NAO', 'nao', 'PAC',
              'pac']
construction = ['Start of Construction/build', 'Complete', 'complete',
                'Tender', 'tender']
operations = ['Full Operations', 'Start of Operation', 'operational', 'Operational',
              'operations', 'Operations', 'operation', 'Operation']
other_gov = ['TAP', 'MPRG', 'Cabinet Office', ' CO ', 'HMT']
consultations = ['Consultation', 'consultation', 'Preferred', 'preferred',
                 'Route', 'route', 'Announcement', 'announcement',
                 'Statutory', 'statutory']
planning = ['DCO', 'dco', 'Planning', 'planning', 'consent', 'Consent',
            'Pre-PIN', 'Pre-OJEU']
ipdc = ['IPDC', 'BICC']

# get filtered milestone chart data
all = m_chart.group_chart(ipdc, start_date, end_date)

#group_chart returns four list
key_names = np.array(all[0])
current_m_data = np.array(all[1])
last_m_data = np.array(all[2])
baseline_m_data = np.array(all[3])

#run the programme
build_charts(key_names,
             current_m_data,
             last_m_data,
             baseline_m_data,
             'IPDC next two years',
             ipdc_date)

#TODO style chart so hides y_axis titles if over a certain number
#TODO style chart to only return project name if all keys are the same.
#TODO improve search for string combos in group_milestone_schedule_data
