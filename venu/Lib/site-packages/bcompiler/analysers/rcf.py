# rcf.py
"""
Analyser to do Reference Class Forecasting on master documents.
"""
import operator
import copy
import datetime
import os
import logging
import re
import sys

from typing import List, Tuple, Dict, Optional

from collections import namedtuple
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series

from ..process.cleansers import DATE_REGEX
from ..utils import ROOT_PATH, runtime_config, CONFIG_FILE

from ..core import Quarter, Master, Row

logger = logging.getLogger('bcompiler.compiler')
runtime_config.read(CONFIG_FILE)

target_master_fn = re.compile(r'^.+_\d{4}.xlsx')

cells_we_want_to_capture = ['Reporting period (GMPP - Snapshot Date)',
                            'Approval MM1',
                            'Approval MM1 Forecast / Actual',
                            'Approval MM3',
                            'Approval MM3 Forecast / Actual',
                            'Approval MM10',
                            'Approval MM10 Forecast / Actual',
                            'Project MM18',
                            'Project MM18 Forecast - Actual',
                            'Project MM19',
                            'Project MM19 Forecast - Actual',
                            'Project MM20',
                            'Project MM20 Forecast - Actual',
                            'Project MM21',
                            'Project MM21 Forecast - Actual']


def _process_masters(path: str) -> Tuple[Quarter, Dict[str, Tuple]]:
    hold = {}
    year = path[-9:][:4]
    quarter = path[-11]
    q = Quarter(int(quarter), int(year))
    m = Master(q, path)
    for p in m.projects:
        pd = m[p]
        hold[p] = pd.pull_keys(cells_we_want_to_capture)
    return q, hold



def create_rcf_output(path: str):
    return _process_masters(path)


def _main_keys(dictionary) -> list:
    return [k for k, _ in dictionary[1].items()]


def _headers(p_name: str, dictionary):
    return [x[0] for x in dictionary[1][p_name]]


def _vals(p_name: str, dictionary):
    return [x[1] for x in dictionary[1][p_name]]


def _inject(lst: list, op, place: int, idxa: int, idxb: int) -> Optional[List]:
    if not isinstance(lst[idxa], datetime.date) and isinstance(lst[idxa], str):
        if re.match(DATE_REGEX, lst[idxa]):
            try:
                ds = lst[idxa].split('-')
                ds = datetime.date(int(ds[0]), int(ds[1]), int(ds[2]))
            except (TypeError, AttributeError):
                logger.warning(f'{lst[idxa]} is not a date so no calculation.')
                return
        else:
            return
    if not isinstance(lst[idxb], datetime.date) and isinstance(lst[idxa], str):
        if re.match(DATE_REGEX, lst[idxb]):
            try:
                ds = lst[idxb].split('-')
                ds = datetime.date(int(ds[0]), int(ds[1]), int(ds[2]))
            except (TypeError, AttributeError):
                logger.warning(f'{lst[idxb]} is not a date so no calculation.')
                return
        else:
            return
    try:
        lst[place] = op(lst[idxa], lst[idxb])
    except TypeError:
        return
    return lst


def _insert_gaps(lst: list, indices: list) -> list:
    for x in indices:
        lst.insert(x, None)
    return lst


def _replace_underscore(name: str):
    return name.replace('/', '_')


def _get_master_files_and_order_them(path: str):
    m = [f for f in os.listdir(path) if re.match(target_master_fn, f)]
    if len(m) == 0:
        raise ValueError("No masters present")
    logger.info(f"Found {m} master files...")
    get_quarter = lambda x: x[-11]
    get_year = lambda x: x[-9::][:4]
    m = sorted(m, key=get_quarter)
    m = sorted(m, key=get_year)
    return m


QueuedWorkbook = namedtuple('QueuedWorkbook', ['project_name', 'file_title', 'workbook'])


def _process_data_cols(worksheet, data_row: list, masters: list, headers: list, start_row: int):

    no_masters = len(masters)

    chart_level_start = 1

    # write the columns used for the chart
    Row(2, start_row, ["SOBC", data_row[3], chart_level_start]).bind(worksheet)
    start_row += no_masters
    chart_level_start += 1

    Row(2, start_row, ["OBC", data_row[6], chart_level_start]).bind(worksheet)
    start_row += no_masters
    chart_level_start += 1

    Row(2, start_row, ["FBC", data_row[9], chart_level_start]).bind(worksheet)
    start_row += no_masters
    chart_level_start += 1

    Row(2, start_row, ["Start of Construction", data_row[15], chart_level_start]).bind(worksheet)
    start_row += no_masters
    chart_level_start += 1

    Row(2, start_row, ["Start of Operation", data_row[18], chart_level_start]).bind(worksheet)
    start_row += no_masters
    chart_level_start += 1

    Row(2, start_row, ["Project End", data_row[21], chart_level_start]).bind(worksheet)
    start_row = start_row + no_masters
    chart_level_start += 1
    return worksheet


def _generate_chart(worksheet, top_row: int, leftmost_col: int) -> ScatterChart:
    chart = ScatterChart()
    chart.title = "RCF"
    chart.style = 13
    chart.height = 18
    chart.width = 28
    chart.x_axis.title = "Days"
    chart.y_axis.title = "Milestone Type"
    xvalues = Reference(worksheet, min_col=3, min_row=10, max_row=33)
    yvalues = Reference(worksheet, min_col=4, min_row=10, max_row=33)
    series = Series(yvalues, xvalues)
    series.marker.size = 6
    chart.series.append(series)
    return chart


def run(output_path: str=None, user_provided_master_path: str=None,) -> None:

    if user_provided_master_path:
        logger.info(f"Using master file location: {user_provided_master_path}")
    else:
        logger.info(f"Using default master location (bcompiler aux directory)")

    if user_provided_master_path is None:
        user_provided_master_path = ROOT_PATH
    if output_path is None:
        output_path = os.path.join(ROOT_PATH, 'output')
    else:
        output_path = output_path

    file_queue: list = []
    flag = False
    try:
        mxs = _get_master_files_and_order_them(user_provided_master_path)
    except ValueError as e:
        logger.critical(f"No masters present in {user_provided_master_path}")
        sys.exit(1)
    chart_data_start_row = 10
    for start_row, f in list(enumerate(mxs, start=2)):
        d = create_rcf_output(os.path.join(user_provided_master_path, f))
        # create a header row first off
        project_titles = _main_keys(d)
        # then take a project at a time
        for proj in project_titles:
            if len(file_queue) > 0:
                for t in file_queue:
                    if proj == t.project_name:
                        wb = t.workbook
                        ws = wb.active
                        flag = True
                        break
                    else:
                        wb = Workbook()
                        ws = wb.active
            else:
                wb = Workbook()
                ws = wb.active
            h_row = _headers(proj, d)
            _insert_gaps(h_row, [3, 6, 9, 12, 15, 18, 21])
            Row(2, 2, h_row).bind(ws)

            d_row = []
            for x in _vals(proj, d):
                d_row.append(x)

            # make spaces in the row
            _insert_gaps(d_row, [3, 6, 9, 12, 15, 18, 21])

            # inject the calculations
            _inject(d_row, operator.sub, 3, 2, 11)
            _inject(d_row, operator.sub, 6, 5, 2)
            _inject(d_row, operator.sub, 9, 8, 5)
            _inject(d_row, operator.sub, 15, 14, 8)
            _inject(d_row, operator.sub, 18, 17, 14)
            _inject(d_row, operator.sub, 21, 20, 17)

            Row(2, start_row + 1, d_row).bind(ws)

            # call process here
            ws = _process_data_cols(ws, d_row, mxs, h_row, chart_data_start_row)
            chart = _generate_chart(ws, 10, 2)
            ws.add_chart(chart, "F10")

            if flag:
                continue
            proj_pack = copy.deepcopy(proj)
#           proj = ''.join([proj, ' ', str(d[0])])
            proj = _replace_underscore(proj)
            proj = proj.replace(' ', '_')
            f_title = f"{proj}_RCF.xlsx"
            file_queue.append(QueuedWorkbook(proj_pack, f_title, wb))
        chart_data_start_row += 1

    try:
        for item in file_queue:
            logger.info(f"Saving {item.file_title} to {output_path}")
            item.workbook.save(os.path.join(output_path, item.file_title))
    except PermissionError:
        logger.critical(
            "Cannot save output file - you already have it open. Close and run again."
        )


if __name__ == '__main__':
    run()
