import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import timedelta
from analysis.data import root_path
import numpy as np

class Baselines:
    def __init__(self, master_data, project_names):
        self.master_data = master_data
        self.project_names = project_names
        self.info = {}
        self.index = {}
        self.get_baseline_data()

    def get_baseline_data(self):
        """
        Given a list of project names in project_names returns
        the two dictionaries baseline_info and baseline_index
        """
        baseline_info = {}
        baseline_index = {}

        for name in self.project_names:
            bc_list = []
            lower_list = []
            for i, master in reversed(list(enumerate(self.master_data))):
                if name in master.projects:
                    approved_bc = master.data[name]['IPDC approval point']
                    quarter = str(master.quarter)
                    if approved_bc not in bc_list:
                        bc_list.append(approved_bc)
                        lower_list.append((approved_bc, quarter, i))
                else:
                    pass
            for i in reversed(range(2)):
                if name in self.master_data[i].projects:
                    approved_bc = self.master_data[i][name]['IPDC approval point']
                    quarter = str(self.master_data[i].quarter)
                    lower_list.append((approved_bc, quarter, i))
                else:
                    quarter = str(self.master_data[i].quarter)
                    lower_list.append((None, quarter, None))

            index_list = []
            for x in lower_list:
                index_list.append(x[2])

            baseline_info[name] = list(reversed(lower_list))
            baseline_index[name] = list(reversed(index_list))

        self.info = baseline_info
        self.index = baseline_index

        #return self.baseline_info, self.baseline_index

class MilestoneData:
    def __init__(self, master_data, baseline_object):
        self.master_data = master_data
        self.baseline = baseline_object
        self.project_dict = {}
        self.group_dict = {}

    def project_data(self, data_to_return):  # renamed to project_data
        """
        Given a list of project names in project_names, and
        integer in data_to_return it
        returns a dictionary containing data for each project.
        """
        # Provided a description of what method does, including the
        # expected parameters and return type. NB, we use " quotes, not '
        # quotes for docstrings like this.

        #Question. Is it possible for this function to return several dicts titled differently above.

        upper_dict = {}

        for name in self.baseline.project_names:
            lower_dict = {}
            raw_list = []
            try:
                p_data = self.master_data[self.baseline.index[name][data_to_return]].data[name]
                for i in range(1, 50):
                    try:
                        try:
                            t = (
                                p_data["Approval MM" + str(i)],
                                p_data["Approval MM" + str(i) + " Forecast / Actual"],
                                p_data["Approval MM" + str(i) + " Notes"],
                            )
                            raw_list.append(t)
                        except KeyError:
                            t = (
                                p_data["Approval MM" + str(i)],
                                p_data["Approval MM" + str(i) + " Forecast - Actual"],
                                p_data["Approval MM" + str(i) + " Notes"],
                            )
                            raw_list.append(t)

                        t = (
                            p_data["Assurance MM" + str(i)],
                            p_data["Assurance MM" + str(i) + " Forecast - Actual"],
                            p_data["Assurance MM" + str(i) + " Notes"],
                        )
                        raw_list.append(t)

                    except KeyError:
                        pass

                for i in range(18, 67):
                    try:
                        t = (
                            p_data["Project MM" + str(i)],
                            p_data["Project MM" + str(i) + " Forecast - Actual"],
                            p_data["Project MM" + str(i) + " Notes"],
                        )
                        raw_list.append(t)
                    except KeyError:
                        pass
            except (KeyError, TypeError):
                pass

            # put the list in chronological order
            sorted_list = sorted(raw_list, key=lambda k: (k[1] is None, k[1]))

            # loop to stop key names being the same. Not ideal as doesn't handle keys that may already have numbers as
            # strings at end of names. But still useful.
            for x in sorted_list:
                if x[0] is not None:
                    if x[0] in lower_dict:
                        for i in range(2, 15):
                            key_name = x[0] + " " + str(i)
                            if key_name in lower_dict:
                                continue
                            else:
                                lower_dict[key_name] = {x[1]: x[2]}
                                break
                    else:
                        lower_dict[x[0]] = {x[1]: x[2]}
                else:
                    pass

            upper_dict[name] = lower_dict

        self.project_dict = upper_dict

        return self.project_dict

    def group_data(self, data_to_return, abbreviations):
        """
        Given a list of project names in project_names,
        returns a dictionary containing data for group of projects
        """

        raw_list = []
        for name in self.baseline.project_names:
            try:
                p_data = self.master_data[self.baseline.index[name][data_to_return]].data[name]
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

        # put the list in chronological order
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

        self.group_dict = output_dict

        return self.group_dict

class MilestoneChartData:
    def __init__(self, current_m_dict, last_m_dict, baseline_m_dict):
        self.current_m_dict = current_m_dict
        self.last_m_dict = last_m_dict
        self.baseline_m_dict = baseline_m_dict
        self.milestone_key_names = []
        self.timedelta_current = []
        self.timedelta_last = []
        self.timedelta_baseline = []

    def group_chart(self, key_of_interest=None,
                    filter_start_date=datetime.date(2000, 1, 1),
                    filter_end_date=datetime.date(2050, 1, 1)):
        """
        Given optional requirements, returns lists containing
        data for a group of project.
        key_of_interest is either default none or a list of strings
        """

        milestone_key_names = []
        timedelta_current = []
        timedelta_last = []
        timedelta_baseline = []

        # all milestone keys and time deltas calculated this way so
        # shown in particular way in output chart
        for m in list(self.current_m_dict.keys()):
            if 'Project - Business Case End Date' in m:  # filter out as dates not helpful
                pass
            else:
                if m is not None:
                    m_d = tuple(self.current_m_dict[m])[0]

                try:
                    if m in list(self.last_m_dict.keys()):
                        m_d_lst = tuple(self.last_m_dict[m])[0]
                    else:
                        m_d_lst = tuple(self.current_m_dict[m])[0]
                except KeyError:
                    m_d_lst = tuple(self.current_m_dict[m])[0]

                if m in list(self.baseline_m_dict.keys()):
                    m_d_bl = tuple(self.baseline_m_dict[m])[0]
                else:
                    m_d_bl = tuple(self.current_m_dict[m])[0]

                if m_d is not None:
                    if filter_start_date <= m_d <= filter_end_date:
                        if key_of_interest is None:
                            milestone_key_names.append(m)

                            timedelta_current.append(m_d)
                            if m_d_lst is not None:
                                timedelta_last.append(m_d_lst)
                            else:
                                timedelta_last.append(m_d)
                            if m_d_bl is not None:
                                timedelta_baseline.append(m_d_bl)
                            else:
                                if m_d_lst is not None:
                                    timedelta_baseline.append(m_d_lst)
                                else:
                                    timedelta_baseline.append(m_d)

                        else:
                            for key in key_of_interest:
                                if key in m:
                                    if m not in milestone_key_names:# prevent repeats
                                        milestone_key_names.append(m)

                                        timedelta_current.append(m_d)
                                        if m_d_lst is not None:
                                            timedelta_last.append(m_d_lst)
                                        else:
                                            timedelta_last.append(m_d)
                                        if m_d_bl is not None:
                                            timedelta_baseline.append(m_d_bl)
                                        else:
                                            if m_d_lst is not None:
                                                timedelta_baseline.append(m_d_lst)
                                            else:
                                                timedelta_baseline.append(m_d)

        self.milestone_key_names = milestone_key_names
        self.timedelta_current = timedelta_current
        self.timedelta_last = timedelta_last
        self.timedelta_baseline = timedelta_baseline

        return self.milestone_key_names, timedelta_current, timedelta_last, timedelta_baseline

class MilestoneCharts:
    def __init__(self, latest_milestone_names, latest_milestone_dates,
                 last_milestone_dates, baseline_milestone_dates, graph_title,
                 ipdc_date):
        self.latest_milestone_names = latest_milestone_names
        self.latest_milestone_dates = latest_milestone_dates
        self.last_milestone_dates = last_milestone_dates
        self.baseline_milestone_dates = baseline_milestone_dates
        self.graph_title = graph_title
        self.ipdc_date = ipdc_date
        #self.milestone_swimlane_charts()
        self.build_charts()

    def milestone_swimlane_charts(self):
        # build scatter chart
        fig, ax1 = plt.subplots()
        fig.suptitle(self.graph_title, fontweight='bold')  # title
        # set fig size
        fig.set_figheight(4)
        fig.set_figwidth(8)

        ax1.scatter(self.baseline_milestone_dates, self.latest_milestone_names, label='Baseline')
        ax1.scatter(self.last_milestone_dates, self.latest_milestone_names, label='Last Qrt')
        ax1.scatter(self.latest_milestone_dates, self.latest_milestone_names, label='Latest Qrt')

        # format the x ticks
        years = mdates.YearLocator()  # every year
        months = mdates.MonthLocator()  # every month
        years_fmt = mdates.DateFormatter('%Y')
        months_fmt = mdates.DateFormatter('%b')

        # calculate the length of the time period covered in chart. Not perfect as baseline dates can distort.
        try:
            td = (self.latest_milestone_dates[-1] - self.latest_milestone_dates[0]).days
            if td <= 365 * 3:
                ax1.xaxis.set_major_locator(years)
                ax1.xaxis.set_minor_locator(months)
                ax1.xaxis.set_major_formatter(years_fmt)
                ax1.xaxis.set_minor_formatter(months_fmt)
                plt.setp(ax1.xaxis.get_minorticklabels(), rotation=45)
                plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45,
                         weight='bold')  # milestone_swimlane_charts(key_name,
                #                           current_m_data,
                #                           last_m_data,
                #                           baseline_m_data,
                #                           'All Milestones')
                # scaling x axis
                # x axis value to no more than three months after last latest milestone date, or three months
                # before first latest milestone date. Hack, can be improved. Text highlights movements off chart.
                x_max = self.latest_milestone_dates[-1] + timedelta(days=90)
                x_min = self.latest_milestone_dates[0] - timedelta(days=90)
                for date in self.baseline_milestone_dates:
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
        except IndexError:  # if milestone dates list is empty:
            pass

        ax1.legend()  # insert legend

        # reverse y axis so order is earliest to oldest
        ax1 = plt.gca()
        ax1.set_ylim(ax1.get_ylim()[::-1])
        ax1.tick_params(axis='y', which='major', labelsize=7)
        ax1.yaxis.grid()  # horizontal lines
        ax1.set_axisbelow(True)
        # ax1.get_yaxis().set_visible(False)

        # for i, txt in enumerate(latest_milestone_names):
        #     ax1.annotate(txt, (i, latest_milestone_dates[i]))

        # Add line of IPDC date, but only if in the time period
        try:
            if self.latest_milestone_dates[0] <= self.ipdc_date <= self.latest_milestone_dates[-1]:
                plt.axvline(self.ipdc_date)
                plt.figtext(0.98, 0.01, 'Line represents when IPDC will discuss Q1 20_21 portfolio management report',
                            horizontalalignment='right', fontsize=6, fontweight='bold')
        except IndexError:
            pass

        # size of chart and fit
        fig.canvas.draw()
        fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

        fig.savefig(root_path / 'output/{}.png'.format(self.graph_title), bbox_inches='tight')

        # plt.close() #automatically closes figure so don't need to do manually.

    def build_charts(self):

        # add \n to y axis labels and cut down if two long
        # labels = ['\n'.join(wrap(l, 40)) for l in latest_milestone_names]
        labels = self.latest_milestone_names
        final_labels = []
        for l in labels:
            if len(l) > 40:
                final_labels.append(l[:35])
            else:
                final_labels.append(l)

        # Chart
        no_milestones = len(self.latest_milestone_names)

        if no_milestones <= 30:
            (np.array(final_labels), np.array(self.latest_milestone_dates),
                              np.array(self.last_milestone_dates),
                              np.array(self.baseline_milestone_dates),
                              self.graph_title, self.ipdc_date)

        if 31 <= no_milestones <= 60:
            half = int(no_milestones / 2)
            MilestoneCharts(np.array(final_labels[:half]),
                                                      np.array(self.latest_milestone_dates[:half]),
                                                      np.array(self.last_milestone_dates[:half]),
                                                      np.array(self.baseline_milestone_dates[:half]),
                                                      self.graph_title, self.ipdc_date)
            title = self.graph_title + ' cont.'
            MilestoneCharts(np.array(final_labels[half:no_milestones]),
                                                      np.array(self.latest_milestone_dates[half:no_milestones]),
                                                      np.array(self.last_milestone_dates[half:no_milestones]),
                                                      np.array(self.baseline_milestone_dates[half:no_milestones]),
                                                      title,
                                                      self.ipdc_date)

        if 61 <= no_milestones <= 90:
            third = int(no_milestones / 3)
            MilestoneCharts(np.array(final_labels[:third]),
                                                      np.array(self.latest_milestone_dates[:third]),
                                                      np.array(self.last_milestone_dates[:third]),
                                                      np.array(self.baseline_milestone_dates[:third]),
                                                      self.graph_title, self.ipdc_date)
            title = self.graph_title + ' cont. 1'
            MilestoneCharts(np.array(final_labels[third:third * 2]),
                                                      np.array(self.latest_milestone_dates[third:third * 2]),
                                                      np.array(self.last_milestone_dates[third:third * 2]),
                                                      np.array(self.baseline_milestone_dates[third:third * 2]),
                                                      title, self.ipdc_date)
            title = self.graph_title + ' cont. 2'
            MilestoneCharts(np.array(final_labels[third * 2:no_milestones]),
                            np.array(self.latest_milestone_dates[third * 2:no_milestones]),
                            np.array(self.last_milestone_dates[third * 2:no_milestones]),
                            np.array(self.baseline_milestone_dates[third * 2:no_milestones]),
                            title, self.ipdc_date)
        pass
