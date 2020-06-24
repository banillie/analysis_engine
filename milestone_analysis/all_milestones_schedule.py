from analysis.data import list_of_masters_all, bc_index, abbreviations, ipdc_date, root_path, hsmrpg
#from analysis.engine_functions import all_milestones_dict
import datetime
from datetime import timedelta
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
from textwrap import wrap


def group_all_milestones_dict(project_names,
                              master_data,
                              baseline_index,
                              data_to_return=int):
    '''
    Function that puts project milestone data in dictionary in order of newest date first.

    Project_names: list of project names of interest / in range
    Master_data: quarter master data set

    Dictionary is structured as {'project name': {'milestone name': datetime.date: 'notes'}}

    '''


    raw_list = []
    for name in project_names:
        try:
            p_data = master_data[baseline_index[name][data_to_return]].data[name]
            for i in range(1, 50):
                try:
                    try:
                        if p_data['Approval MM' + str(i)] is None:
                            pass
                        else:
                            key_name = abbreviations[name] + ', ' + p_data['Approval MM' + str(i)]
                            t = (key_name,
                                 p_data['Approval MM' + str(i) + ' Forecast / Actual'],
                                 p_data['Approval MM' + str(i) + ' Notes'])
                            raw_list.append(t)
                    except KeyError:
                        if p_data['Approval MM' + str(i)] is None:
                            pass
                        else:
                            key_name = abbreviations[name] + ', ' + p_data['Approval MM' + str(i)]
                            t = (key_name,
                                 p_data['Approval MM' + str(i) + ' Forecast - Actual'],
                                 p_data['Approval MM' + str(i) + ' Notes'])
                            raw_list.append(t)

                    if p_data['Assurance MM' + str(i)] is None:
                        pass
                    else:
                        key_name = abbreviations[name] + ', ' + p_data['Assurance MM' + str(i)]
                        t = (key_name,
                             p_data['Assurance MM' + str(i) + ' Forecast - Actual'],
                             p_data['Assurance MM' + str(i) + ' Notes'])
                        raw_list.append(t)

                except KeyError:
                    pass

            for i in range(18, 67):
                try:
                    if p_data['Project MM' + str(i)] is None:
                        pass
                    else:
                        key_name = abbreviations[name] + ', ' + p_data['Project MM' + str(i)]
                        t = (key_name,
                             p_data['Project MM' + str(i) + ' Forecast - Actual'],
                             p_data['Project MM' + str(i) + ' Notes'])
                        raw_list.append(t)
                except KeyError:
                    pass
        except (KeyError, TypeError):
            pass

    #put the list in chronological order
    sorted_list = sorted(raw_list, key=lambda k: (k[1] is None, k[1]))

    # loop to stop key names being the same. Not ideal as doesn't handle keys that may already have numbers as
    # strings at end of names. But still useful.
    output_dict = {}
    for x in sorted_list:
        if x[0] is not None:
            if x[0] in output_dict:
                for i in range(2, 15):
                    key_name = x[0] + ' ' + str(i)
                    if key_name in output_dict:
                        continue
                    else:
                        output_dict[key_name] = {x[1]: x[2]}
                        break
            else:
                output_dict[x[0]] = {x[1]: x[2]}
        else:
            pass

    return output_dict

def group_milestone_schedule_data(latest_m_dict,
                                  last_m_dict,
                                  baseline_m_dict,
                                  key_of_interest=None,
                                  filter_start_date=datetime.date(2000, 1, 1),
                                  filter_end_date=datetime.date(2050, 1, 1)):

    milestone_names = []
    mile_d_l_lst = []
    mile_d_last_lst = []
    mile_d_bl_lst = []

    #lengthy for loop designed so that all milestones and dates are stored and shown in output chart, even if they
    #were not present in last and baseline data reporting
    for m in list(latest_m_dict.keys()):
        if 'Project - Business Case End Date' in m:  # filter out as dates not helpful
            pass
        else:
            if m is not None:
                m_d = tuple(latest_m_dict[m])[0]

            try:
                if m in list(last_m_dict.keys()):
                    m_d_lst = tuple(last_m_dict[m])[0]
                else:
                    m_d_lst = tuple(latest_m_dict[m])[0]
            except KeyError:
                m_d_lst = tuple(latest_m_dict[m])[0]

            if m in list(baseline_m_dict.keys()):
                m_d_bl = tuple(baseline_m_dict[m])[0]
            else:
                m_d_bl = tuple(latest_m_dict[m])[0]

            if m_d is not None:
                if filter_start_date <= m_d <= filter_end_date:
                    if key_of_interest is None:
                        milestone_names.append(m)

                        mile_d_l_lst.append(m_d)
                        if m_d_lst is not None:
                            mile_d_last_lst.append(m_d_lst)
                        else:
                            mile_d_last_lst.append(m_d)
                        if m_d_bl is not None:
                            mile_d_bl_lst.append(m_d_bl)
                        else:
                            if m_d_lst is not None:
                                mile_d_bl_lst.append(m_d_lst)
                            else:
                                mile_d_bl_lst.append(m_d)

                    else:
                        if key_of_interest in m:
                            milestone_names.append(m)

                            mile_d_l_lst.append(m_d)
                            if m_d_lst is not None:
                                mile_d_last_lst.append(m_d_lst)
                            else:
                                mile_d_last_lst.append(m_d)
                            if m_d_bl is not None:
                                mile_d_bl_lst.append(m_d_bl)
                            else:
                                if m_d_lst is not None:
                                    mile_d_bl_lst.append(m_d_lst)
                                else:
                                    mile_d_bl_lst.append(m_d)

    return milestone_names, mile_d_l_lst, mile_d_last_lst, mile_d_bl_lst


def milestone_swimlane_charts(latest_milestone_names,
                              latest_milestone_dates,
                              last_milestone_dates,
                              baseline_milestone_dates,
                              graph_title):


    #build scatter chart
    fig, ax1 = plt.subplots()
    fig.suptitle(graph_title, fontweight='bold')  # title
    # set fig size
    # fig.set_figheight(6)
    # fig.set_figwidth(8)

    ax1.scatter(baseline_milestone_dates, latest_milestone_names, label='Baseline')
    ax1.scatter(last_milestone_dates, latest_milestone_names, label='Last Qrt')
    ax1.scatter(latest_milestone_dates, latest_milestone_names, label='Latest Qrt')

    # format the x ticks
    years = mdates.YearLocator()  # every year
    months = mdates.MonthLocator()  # every month
    years_fmt = mdates.DateFormatter('%Y')
    months_fmt = mdates.DateFormatter('%b')

    # calculate the length of the time period covered in chart. Not perfect as baseline dates can distort.
    try:
        td = (latest_milestone_dates[-1] - latest_milestone_dates[0]).days
        if td <= 365*3:
            ax1.xaxis.set_major_locator(years)
            ax1.xaxis.set_minor_locator(months)
            ax1.xaxis.set_major_formatter(years_fmt)
            ax1.xaxis.set_minor_formatter(months_fmt)
            plt.setp(ax1.xaxis.get_minorticklabels(), rotation=45)
            plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, weight='bold')
            # scaling x axis
            # x axis value to no more than three months after last latest milestone date, or three months
            # before first latest milestone date. Hack, can be improved. Text highlights movements off chart.
            x_max = last_milestone_dates[-1] + timedelta(days=90)
            x_min = last_milestone_dates[0] - timedelta(days=90)
            for date in baseline_milestone_dates:
                if date > x_max:
                    ax1.set_xlim(x_min, x_max)
                    plt.figtext(0.98, 0.03,
                                'Check full schedule to see all milestone movements',
                                horizontalalignment='right', fontsize=6, fontweight='bold')
                if date < x_min:
                    ax1.set_xlim(x_min, x_max)
                    plt.figtext(0.98, 0.03,
                                'Check full schedule to see all milestone movements',
                                horizontalalignment='right', fontsize=6, fontweight='bold')
        else:
            ax1.xaxis.set_major_locator(years)
            ax1.xaxis.set_minor_locator(months)
            ax1.xaxis.set_major_formatter(years_fmt)
            plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, weight='bold')
    except IndexError: #if milestone dates list is empty:
        pass

    ax1.legend() #insert legend

    #reverse y axis so order is earliest to oldest
    ax1 = plt.gca()
    ax1.set_ylim(ax1.get_ylim()[::-1])
    #ax1.tick_params(axis='y', which='major', labelsize=7)
    ax1.get_yaxis().set_visible(False)

    #Add line of IPDC date, but only if in the time period
    try:
        if latest_milestone_dates[0] <= ipdc_date <= latest_milestone_dates[-1]:
            plt.axvline(ipdc_date)
            plt.figtext(0.98, 0.01, 'Line represents when IPDC will discuss Q1 20_21 portfolio management report',
                        horizontalalignment='right', fontsize=6, fontweight='bold')
    except IndexError:
        pass

    #size of chart and fit
    fig.canvas.draw()
    fig.tight_layout(rect=[0, 0.03, 1, 0.95]) #for title

    fig.savefig(root_path/'output/{}.png'.format(graph_title), bbox_inches='tight')
    #plt.close() #automatically closes figure so don't need to do manually.

    #os.remove('schedule.png')

def build_charts(latest_milestone_names,
                 latest_milestone_dates,
                 last_milestone_dates,
                 baseline_milestone_dates,
                 graph_title):
    pass


p_n_list = list_of_masters_all[0].projects
current_m = group_all_milestones_dict(p_n_list, list_of_masters_all, bc_index, 0)
last_m = group_all_milestones_dict(p_n_list, list_of_masters_all, bc_index, 1)
baseline_m = group_all_milestones_dict(p_n_list, list_of_masters_all, bc_index, 2)

chart_data = group_milestone_schedule_data(current_m, last_m, baseline_m)

key_name = np.array(chart_data[0])
current_m_data = np.array(chart_data[1])
last_m_data = np.array(chart_data[2])
baseline_m_data = np.array(chart_data[3])

milestone_swimlane_charts(key_name,
                          current_m_data,
                          last_m_data,
                          baseline_m_data,
                          'All Milestones')

#TODO style chart so hides y_axis titles if over a certain number
#TODO style chart to only return project name if all keys are the same.
#TODO improve search for string combos in group_milestone_schedule_data
#TODO explore whether matplotlib output file format can be improved