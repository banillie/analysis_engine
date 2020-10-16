from data_mgmt.data import MilestoneData, MilestoneChartData, \
    Master, Projects, master_data_list, root_path, blue_line_date, \
    abbreviations, CombinedData
import datetime
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import timedelta
from openpyxl import load_workbook


def milestone_swimlane_charts(latest_milestone_names,
                              latest_milestone_dates,
                              last_milestone_dates,
                              baseline_milestone_dates,
                              baseline_milestone_dates_two,
                              graph_title,
                              analysis_date):

    # build scatter chart
    fig, ax1 = plt.subplots()
    fig.suptitle(graph_title, fontweight='bold')  # title

    # set fig size
    fig.set_figheight(4)
    fig.set_figwidth(8)

    #ax1.scatter(baseline_milestone_dates_two, latest_milestone_names, label='Baseline (last)')
    ax1.scatter(baseline_milestone_dates, latest_milestone_names, label='Baseline (current)')
    # ax1.scatter(last_milestone_dates, latest_milestone_names, label='Last Qrt')
    ax1.scatter(latest_milestone_dates, latest_milestone_names, label='Latest/Achieved')

    # format the x ticks
    years = mdates.YearLocator()  # every year
    months = mdates.MonthLocator()  # every month
    years_fmt = mdates.DateFormatter('%Y')
    months_fmt = mdates.DateFormatter('%b')
    ax1.xaxis.set_major_locator(years)
    ax1.xaxis.set_minor_locator(months)
    ax1.xaxis.set_major_formatter(years_fmt)
    ax1.xaxis.set_minor_formatter(months_fmt)
    plt.setp(ax1.xaxis.get_minorticklabels(), rotation=45)
    plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, weight='bold')


    # """calculate the length of the time period covered in chart.
    # Not perfect as baseline dates can distort."""
    # try:
    #     td = (latest_milestone_dates[-1] - latest_milestone_dates[0]).days
    #     if td >= 365 * 3:
    #         print('Yes')
    #         ax1.xaxis.set_major_locator(years)
    #         ax1.xaxis.set_minor_locator(months)
    #         ax1.xaxis.set_major_formatter(years_fmt)
    #         ax1.xaxis.set_minor_formatter(months_fmt)
    #         plt.setp(ax1.xaxis.get_minorticklabels(), rotation=45)
    #         plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, weight='bold')
    #
    #         # scaling x axis
    #         # x axis value to no more than three months after last latest milestone date, or three months
    #         # before first latest milestone date. Hack, can be improved. Text highlights movements off chart.
    #         x_max = latest_milestone_dates[-1] + timedelta(days=90)
    #         x_min = latest_milestone_dates[0] - timedelta(days=90)
    #         for date in baseline_milestone_dates:
    #             if date > x_max:
    #                 ax1.set_xlim(x_min, x_max)
    #                 plt.figtext(0.98, 0.03,
    #                             'Check full schedule to see all milestone movements',
    #                             horizontalalignment='right', fontsize=6, fontweight='bold')
    #             if date < x_min:
    #                 ax1.set_xlim(x_min, x_max)
    #                 plt.figtext(0.98, 0.03,
    #                             'Check full schedule to see all milestone movements',
    #                             horizontalalignment='right', fontsize=6, fontweight='bold')
    #     else:
    #         ax1.xaxis.set_major_locator(years)
    #         ax1.xaxis.set_minor_locator(months)
    #         ax1.xaxis.set_major_formatter(years_fmt)
    #         ax1.xaxis.set_minor_formatter(months_fmt)
    #         plt.setp(ax1.xaxis.get_minorticklabels(), rotation=45)
    #         plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, weight='bold')
    #
    #
    # except IndexError:  # if milestone dates list is empty:
    #     pass

    ax1.legend()  # insert legend

    # reverse y axis so order is earliest to oldest
    ax1 = plt.gca()
    ax1.set_ylim(ax1.get_ylim()[::-1])
    ax1.tick_params(axis='y', which='major', labelsize=7)
    ax1.yaxis.grid()  # horizontal lines
    ax1.set_axisbelow(True)
    # ax1.get_yaxis().set_visible(False)

    # Add line of analysis date, but only if in the time period
    try:
        if latest_milestone_dates[0] <= analysis_date <= latest_milestone_dates[-1]:
            plt.axvline(analysis_date)
            plt.figtext(0.98, 0.01, 'Line represents date analysis compiled',
                        horizontalalignment='right', fontsize=6, fontweight='bold')
            # plt.figtext(0.98, 0.01, 'Line represents when IPDC will discuss Q1 20_21 portfolio management report',
            #            horizontalalignment='right', fontsize=6, fontweight='bold')
    except IndexError:
        pass

    # size of chart and fit
    fig.canvas.draw()
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

    fig.savefig(root_path / 'output/{}.png'.format(graph_title), bbox_inches='tight')

    # plt.close() #automatically closes figure so don't need to do manually.


def build_charts(latest_milestone_names,
                 latest_milestone_dates,
                 last_milestone_dates,
                 baseline_milestone_dates,
                 baseline_milestone_dates_two,
                 graph_title,
                 ipdc_date,
                 no_of_labels):
    """
    calculates how many graphical outputs to produced
    based on number of milestones. Milestone key names,
    dates, graph title, date for blue line to represent,
    and number of labels to have on each graph
    are passed in.
    """

    # axis labels are reduced if two long
    labels = latest_milestone_names
    final_labels = []
    for l in labels:
        if len(l) >= 40:
            final_labels.append(l[:40])
        else:
            final_labels.append(l)

    #  Charts are built
    no_milestones = len(latest_milestone_names)

    if no_milestones <= no_of_labels:
        milestone_swimlane_charts(np.array(final_labels), np.array(latest_milestone_dates),
                                  np.array(last_milestone_dates),
                                  np.array(baseline_milestone_dates),
                                  np.array(baseline_milestone_dates_two),
                                  graph_title, ipdc_date)

    if no_of_labels + 1 <= no_milestones <= no_of_labels*2:
        half = int(no_milestones / 2)
        milestone_swimlane_charts(np.array(final_labels[:half]),
                                  np.array(latest_milestone_dates[:half]),
                                  np.array(last_milestone_dates[:half]),
                                  np.array(baseline_milestone_dates[:half]),
                                  np.array(baseline_milestone_dates_two[:half]),
                                  graph_title, ipdc_date)
        title = graph_title + ' cont.'
        milestone_swimlane_charts(np.array(final_labels[half:no_milestones]),
                                  np.array(latest_milestone_dates[half:no_milestones]),
                                  np.array(last_milestone_dates[half:no_milestones]),
                                  np.array(baseline_milestone_dates[half:no_milestones]),
                                  np.array(baseline_milestone_dates_two[half:no_milestones]),
                                  title, ipdc_date)

    if (no_of_labels*2) + 1 <= no_milestones <= no_of_labels*3:
        third = int(no_milestones / 3)
        milestone_swimlane_charts(np.array(final_labels[:third]),
                                  np.array(latest_milestone_dates[:third]),
                                  np.array(last_milestone_dates[:third]),
                                  np.array(baseline_milestone_dates[:third]),
                                  np.array(baseline_milestone_dates_two[:third]),
                                  graph_title, ipdc_date)
        title = graph_title + ' cont. 1'
        milestone_swimlane_charts(np.array(final_labels[third:third * 2]),
                                  np.array(latest_milestone_dates[third:third * 2]),
                                  np.array(last_milestone_dates[third:third * 2]),
                                  np.array(baseline_milestone_dates[third:third * 2]),
                                  np.array(baseline_milestone_dates_two[third:third * 2]),
                                  title, ipdc_date)
        title = graph_title + ' cont. 2'
        milestone_swimlane_charts(np.array(final_labels[third * 2:no_milestones]),
                                  np.array(latest_milestone_dates[third * 2:no_milestones]),
                                  np.array(last_milestone_dates[third * 2:no_milestones]),
                                  np.array(baseline_milestone_dates[third * 2:no_milestones]),
                                  np.array(baseline_milestone_dates_two[third * 2:no_milestones]),
                                  title, ipdc_date)
    pass


"""Get data"""
mst = Master(master_data_list[1:], Projects.hsmrpg)  # get master data and specify projects
mst.baseline_data('Re-baseline IPDC milestones')  # get baseline information of interest
milestone_data = MilestoneData(mst, abbreviations)  # get milestone data

hsmrpg_milestone_wb = load_workbook("/home/will/Documents/analysis_engine/input/exported_milestones_HSMRPG.xlsx")
combined_milestone_data = CombinedData(hsmrpg_milestone_wb, milestone_data)

"""filtering data options for the chart."""
# Format year, month, day
start_date = datetime.date(2020, 9, 1)
end_date = datetime.date(2021, 9, 1)

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
                 'Statutory', 'statutory', 'PRA']
planning = ['DCO', 'dco', 'Planning', 'planning', 'consent', 'Consent',
            'Pre-PIN', 'Pre-OJEU', 'Initiation', 'initiation']
ipdc = ['IPDC', 'BICC']
he_search = ['Start of Construction/build', 'DCO', 'dco', 'PRA',
             'Preferred', 'preferred', 'Route', 'route',
             'Annoucement', 'announcement', 'submission',
             'PVR'
             'Submission']
remove = ['Benefits']

"""Process data into format for the chart"""
mcd = MilestoneChartData(milestone_data_object=combined_milestone_data,
                         keys_of_interest=None,
                         keys_not_of_interest=None,
                         filter_start_date=start_date,
                         filter_end_date=end_date)

build_charts(mcd.group_keys,
             mcd.group_current_tds,
             mcd.group_last_tds,
             mcd.group_baseline_tds,
             mcd.group_baseline_tds_two,
             'HSMRPG schedule',
             blue_line_date,
             20)



