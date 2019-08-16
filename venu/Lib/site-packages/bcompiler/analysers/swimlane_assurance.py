import collections
import datetime
import os
from typing import Tuple

import openpyxl
from openpyxl.chart import ScatterChart, Reference, Series
# typing imports

from .utils import MASTER_XLSX, logger, projects_in_master, diff_date_list, date_convertor
from ..utils import ROOT_PATH, runtime_config, CONFIG_FILE

runtime_config.read(CONFIG_FILE)

DAY_RANGE = int(runtime_config['AnalyserSwimlaneAssurance']['day_range'])
BLOCK_START = int(runtime_config['AnalyserSwimlaneAssurance']['block_start'])
BLOCK_SKIP = int(runtime_config['AnalyserSwimlaneAssurance']['block_skip'])
BLOCK_END = int(runtime_config['AnalyserSwimlaneAssurance']['block_end'])
FORECAST_ACTUAL_SKIP = int(runtime_config['AnalyserSwimlaneAssurance']['forecast_actual_skip'])
MILESTONES_TO_COLLECT = int(runtime_config['AnalyserSwimlaneAssurance']['milestones_to_collect'])
CHART_ANCHOR_CELL = runtime_config['AnalyserSwimlaneAssurance']['chart_anchor_cell']
CHART_TITLE = runtime_config['AnalyserSwimlaneAssurance']['chart_title']
CHART_X_AXIS_TITLE = runtime_config['AnalyserSwimlaneAssurance']['chart_x_axis_title']
CHART_Y_AXIS_TITLE = runtime_config['AnalyserSwimlaneAssurance']['chart_y_axis_title']
CHART_HEIGHT = int(runtime_config['AnalyserSwimlaneAssurance']['chart_height'])
CHART_WIDTH = int(runtime_config['AnalyserSwimlaneAssurance']['chart_width'])
CHART_STYLE = int(runtime_config['AnalyserSwimlaneAssurance']['chart_style'])
CHART_X_AXIS_MAJOR_UNIT = int(runtime_config['AnalyserSwimlaneAssurance']['chart_x_axis_major_unit'])
CHART_Y_AXIS_MAJOR_UNIT = int(runtime_config['AnalyserSwimlaneAssurance']['chart_y_axis_major_unit'])

if runtime_config['AnalyserSwimlane']['grey_markers'] in ['True', 'true', 'yes', 'on']:
    GREYMARKER = True
else:
    GREYMARKER = False

_marker_colours = [
    "FF0000",
    "a86001",
    "4401a8",
    "a801a5",
    "016da8",
    "01a852",
    "FF0000",
]

_grey_marker_colours = ["969696"] * 7


def date_range_milestones(source_sheet, output_sheet, cols: tuple,
                          start_row: int, column: int, date_ends: list):
    """
    Helper function to populate Column B in resulting milestones spreadsheet.
    Uses start and end dates to define boundaries to milestones.
    """
    base_date = date_ends[0]
    current_row = start_row
    dates = diff_date_list(*date_ends)
    for i in range(*cols):
        time_line_date = source_sheet.cell(row=i, column=column).value
        time_line_date = date_convertor(time_line_date)
        try:
            if time_line_date in dates:
                output_sheet.cell(
                    row=current_row,
                    column=3,
                    value=(time_line_date - base_date).days)
                logger.debug(f"Using date range: written {time_line_date} to "
                             f"row: {current_row} col: {column}")
        except TypeError:
            pass
        finally:
            current_row += 1
    return output_sheet


def date_diff_column(source_sheet, output_sheet, cols: tuple, start_row: int, column: int,
                     interested_range: int):
    """Helper function to populate Column B in the resulting milestones spreadsheet."""
    today = datetime.date.today()
    current_row = start_row
    for i in range(*cols):
        time_line_date = source_sheet.cell(row=i, column=column).value
        time_line_date = date_convertor(time_line_date)
        try:
            difference = (time_line_date - today).days
            if difference in range(1, interested_range):
                output_sheet.cell(row=current_row, column=3, value=difference)
                logger.debug(f"Not using date range: written {time_line_date} to "
                             f"row: {current_row} col: {column}")
        except TypeError:
            pass
        finally:
            current_row += 1
    return output_sheet


def splat_date_range(dt: str):
    """Helper function to parse a date in dd/mm/yy format to a list of ints."""
    xs = dt.split('/')
    if len(xs[-1]) == 2:
        xs[-1] = "".join(["20", xs[-1]])
        logger.debug(f"Handling two digit date in argument. Assuming year is {xs[-1]}.")
    xs = [xs[2], xs[1], xs[0]]
    logger.debug(f"Splatting {dt}")
    return [int(i) for i in xs]


def gather_data(start_row: int,
                project_number: int,
                newwb: openpyxl.Workbook,
                block_start_row: int = BLOCK_START,
                interested_range: int = DAY_RANGE,
                master_path=None,
                date_range=None):
    """
    Gather data from
    :type int: start_row
    :type int project_number
    :type openpyxl.Workbook: newwb
    :type int: block_start_row
    :type int: interested_range
    :rtype: Tuple
    """
    newsheet: Worksheet = newwb.active
    col = project_number + 1
    start_row = start_row + 1

    if master_path:
        master = master_path
        logger.debug(f"Using master path: {master_path}")
    else:
        master = MASTER_XLSX
        logger.debug(f"Using master path: {master}")

    wb = openpyxl.load_workbook(master)
    sheet = wb.active

    # print project title
    newsheet.cell(
        row=start_row - 1, column=1, value=sheet.cell(row=1, column=col).value)
    logger.info(f"Processing: {sheet.cell(row=1, column=col).value}")

    x = start_row
    for i in range(block_start_row, BLOCK_END + 2, BLOCK_SKIP):
        val = sheet.cell(row=i, column=col).value
        newsheet.cell(row=x, column=1, value=val)
        logger.debug(f"Writing {val} to row: {x} col: 1")
        x += 1
    x = start_row
    for i in range(block_start_row + FORECAST_ACTUAL_SKIP, BLOCK_END + 2, BLOCK_SKIP):
        val = sheet.cell(row=i, column=col).value
        if isinstance(val, datetime.datetime):
            val = val.date()
        newsheet.cell(row=x, column=2, value=val)
        logger.debug(f"Writing {val} to row: {x} col: 2")
        x += 1

    # process the sheet to populate Column B
    if date_range:
        newsheet = date_range_milestones(
            sheet, newsheet, (BLOCK_START + 3, BLOCK_END + 2, BLOCK_SKIP), start_row, col,
            [datetime.date(*splat_date_range(date_range[0])),
             datetime.date(*splat_date_range(date_range[1]))])
    else:
        newsheet = date_diff_column(sheet, newsheet, (BLOCK_START + 3, BLOCK_END + 2, BLOCK_SKIP), start_row, col,
                                    interested_range)

    for i in range(start_row, start_row + MILESTONES_TO_COLLECT):
        newsheet.cell(row=i, column=4, value=project_number)
        logger.debug(f"Writing {project_number} to row: {i} col: 4")

    return newwb, start_row




def _segment_series() -> Tuple:
    """Generator for step value when stepping through rows within a project block."""
    cut = dict(pvr_gate_zero=2, sobc=1, obc=1, fbc=1, readiness_closure_exit=3, the_rest=10)
    for item in cut.items():
        yield item


def _series_producer(sheet, start_row: int, step: int) -> Tuple[Series, int]:
    """
    Generates a single Series() object, containing a Reference() object for x and y values for the chart.
    Implemented as part of a loop; also returns new_start which is the row number it should continue with
    on the next loop.
    :type sheet: Worksheet
    :type start_row: int
    :type step: int
    :return: tuple of items from cut
    """
    xvalues = Reference(
        sheet, min_col=3, min_row=start_row, max_row=start_row + step)
    values = Reference(
        sheet, min_col=4, min_row=start_row, max_row=start_row + step)
    series = Series(values, xvalues)
    new_start = start_row + step + 1
    return series, new_start


def _row_calc(project_number: int) -> Tuple[int, int]:
    """
    Helper function to calculate row numbers when writing column of project values to cols A & B.
    :type project_number: int
    :return:  tuple of form (project_number, calculated rows in project block)
    """
    if project_number == 1:
        return 1, 1
    if project_number == 2:
        return 2, 20
    else:
        return (project_number,
                (project_number + MILESTONES_TO_COLLECT) + ((project_number - 2) * MILESTONES_TO_COLLECT))


def run(output_path=None, user_provided_master_path=None, date_range=None):
    """
    Main function to run this analyser.
    :param output_path:
    :param user_provided_master_path:
    :return:
    """

    if user_provided_master_path:
        logger.info(f"Using master file: {user_provided_master_path}")
        NUMBER_OF_PROJECTS = projects_in_master(user_provided_master_path)
    else:
        logger.info(f"Using default master file (refer to config.ini)")
        NUMBER_OF_PROJECTS = projects_in_master(
            os.path.join(ROOT_PATH,
                         runtime_config['MasterForAnalysis']['name']))

    wb = openpyxl.Workbook()
    segment_series_generator = _segment_series()

    logger.debug(f"Using block_start of {BLOCK_START}")
    logger.debug(f"Using day_range of {DAY_RANGE}")
    logger.debug(f"Using block_skip of {BLOCK_SKIP}")
    logger.debug(f"Using block_end of {BLOCK_END}")
    logger.debug(f"Using forecast_actual_skip of {FORECAST_ACTUAL_SKIP}")

    for p in range(1, NUMBER_OF_PROJECTS + 1):
        proj_num, st_row = _row_calc(p)
        wb = gather_data(
            st_row,
            proj_num,
            wb,
            block_start_row=BLOCK_START,
            interested_range=DAY_RANGE,
            master_path=user_provided_master_path,
            date_range=date_range)[0]

    chart = ScatterChart()
    chart.title = CHART_TITLE
    chart.style = CHART_STYLE
    chart.height = CHART_HEIGHT
    chart.width = CHART_WIDTH
    chart.x_axis.title = CHART_X_AXIS_TITLE
    chart.y_axis.title = CHART_Y_AXIS_TITLE
    chart.legend = None
    chart.x_axis.majorUnit = CHART_X_AXIS_MAJOR_UNIT
    chart.x_axis.minorGridlines = None
    chart.y_axis.majorUnit = CHART_Y_AXIS_MAJOR_UNIT

    derived_end = 2

    if GREYMARKER:
        markercol = _grey_marker_colours
    else:
        markercol = _marker_colours

    for p in range(NUMBER_OF_PROJECTS):
        for i in range(
                1, 7
        ):  # 7 here is hard-coded number of segments within a project series (ref: dict in _segment_series()
            if i == 1:
                inner_start_row = derived_end
            else:
                inner_start_row = derived_end
            _inner_step = next(segment_series_generator)
            series, derived_end = _series_producer(wb.active, inner_start_row,
                                                   _inner_step[1] - 1)
            if _inner_step[0] == 'pvr_gate_zero':
                series.marker.symbol = "diamond"
                series.marker.graphicalProperties.solidFill = markercol[0]
            elif _inner_step[0] == 'sobc':
                series.marker.symbol = "circle"
                series.marker.graphicalProperties.solidFill = markercol[0]
            elif _inner_step[0] == 'obc':
                series.marker.symbol = "triangle"
                series.marker.graphicalProperties.solidFill = markercol[0]
            elif _inner_step[0] == 'fbc':
                series.marker.symbol = "square"
                series.marker.graphicalProperties.solidFill = markercol[0]
            elif _inner_step[0] == 'readiness_closure_exit':
                series.marker.symbol = "plus"
                series.marker.graphicalProperties.solidFill = markercol[0]
            else:
                series.marker.symbol = "triangle"
                series.marker.graphicalProperties.solidFill = markercol[0]
            series.marker.size = 10
            chart.series.append(series)
        segment_series_generator = _segment_series()
        derived_end = derived_end + 1

    wb.active.add_chart(chart, CHART_ANCHOR_CELL)
    try:
        if output_path:
            wb.save(os.path.join(output_path[0], 'swimlane_assurance_milestones.xlsx'))
            logger.info(f"Saved swimlane_assurance_milestones.xlsx to {output_path}")
        else:
            output_path = os.path.join(ROOT_PATH, 'output')
            wb.save(os.path.join(output_path, 'swimlane_assurance_milestones.xlsx'))
            logger.info(f"Saved swimlane_assurance_milestones.xlsx to {output_path}")
    except PermissionError:
        logger.critical(
            "Cannot save swimlane_assurance_milestones.xlsx file - you already have it open. Close and run again."
        )
        return


if __name__ == "__main__":
    run()
