import collections
import datetime
import logging
import os
import re
import sys

from typing import List, Tuple

from openpyxl import load_workbook

from bcompiler.utils import ROOT_PATH, runtime_config

MASTER_XLSX = os.path.join(ROOT_PATH, runtime_config['MasterForAnalysis']['name'])
logger = logging.getLogger('bcompiler.compiler')


def list_of_milestone_type(project_dict: dict, milestone_type: str) -> List[Tuple]:
    return [item for item in list(iter(project_dict.items())) if item[0].startswith(f'{milestone_type} MM')]


DatamapLine = collections.namedtuple('DatamapLine', ['key', 'sheet'])


class DatamapData:
    # REPLACE THIS WITH DATA FROM THE DATAMAP CSV
    d = """Project/Programme Name,Summary,C6
    SRO First Name,Summary,C5
    SRO Last Name,Summary,C6
    DRO Last Address,Summary,C7"""

    # CHANGE THESE LINES ACCORDINGLY
    lines = d.split('\n')
    keys = [item.split(',')[0] for item in lines]
    sheets = [item.split(',')[1] for item in lines]
    refs = [item.split(',')[2] for item in lines]

    def __init__(self):
        self._lines = [DatamapLine(key.lstrip(), sheet) for key, sheet in zip(self.keys, self.sheets)]

    def __len__(self):
        return len(self._lines)

    def __getitem__(self, position):
        return self._lines[position]


# IMPLEMENT A CLASS THAT SEEKS OUT ALL THE APPROVAL AND ASSURANCE MILESTONES
# CELLS WE NEED AND MAKE THEM AN ITERATOR



def date_convertor(date_thing):
    """
    Date thing could be a datetime object, date object or a sting that looks like
    a date. Our job is to ensure it leaves here as a date object, or return an exception.
    """
    day_first_regex = r"^(\d{1,2})(/|-)(\d{1,2})(/|-)(\d{2,4})"
    year_first_regex = r"^(\d{2,4})(/|-)(\d{1,2})(/|-)(\d{1,2})"

    if not isinstance(date_thing, (str, datetime.datetime, datetime.date)):
        if date_thing is not None:
            logger.warning(f"{date_thing} isn't a date so not handling")
            return date_thing
        else:
            return date_thing
    if isinstance(date_thing, datetime.datetime):
        return date_thing.date()
    df = re.match(day_first_regex, date_thing)
    yf = re.match(year_first_regex, date_thing)
    if df:
        try:
            return datetime.date(int(df.group(5)), int(df.group(3)), int(df.group(1)))
        except ValueError:
            if date_thing is not None:
                logger.warning(f"{date_thing} does not appear to be a valid date.")
            return date_thing
    if yf:
        try:
            return datetime.date(int(yf.group(1)), int(yf.group(3)), int(yf.group(5)))
        except ValueError:
            if date_thing is not None:
                logger.warning(f"{date_thing} does not appear to be a valid date.")
            return date_thing


def diff_date_list(start_date: datetime.date, end_date: datetime.date) -> list:
    """
    Return a list of date objects given start and end date objects.
    """
    return [end_date - datetime.timedelta(days=x) for x in range(0, (end_date - start_date).days)]


def get_number_of_projects(source_wb) -> int:
    """
    Simple helper function to get an accurate number of projects in a master.
    Also strips out any additional columns that openpyxl thinks exist actively
    in the spreadsheet.

    Returns an integer.
    """
    ws = source_wb.active
    top_row = next(ws.rows)  # ws.rows produces a "generator"; use next() to get next value
    top_row = list(top_row)[1:]  # we don't want the first column value
    top_row = [i.value for i in top_row if i.value is not None]  # list comprehension to remove None values
    return len(top_row)


def project_titles_in_master(master: str):
    try:
        wb = load_workbook(master)
    except FileNotFoundError:
        logger.critical("Please ensure you specify a master file in the command or use the correctly named"
                        " master file in your auxiliary directory.")
        sys.exit(1)
    ws = wb.active
    top_row = next(ws.rows)
    top_row = list(top_row)[1:]
    top_row = [i.value for i in top_row if i.value is not None]
    return top_row


def projects_in_master(master: int):
    """
    Return count of items in list of project titles in master.
    :type str: master
    :return int: count of list of projects
    """
    try:
        wb = load_workbook(master)
    except FileNotFoundError:
        logger.critical("Please ensure you specify a master file in the command or use the correctly named"
                        " master file in your auxiliary directory.")
        sys.exit(1)
    ws = wb.active
    top_row = next(ws.rows)
    top_row = list(top_row)[1:]
    top_row = [i.value for i in top_row if i.value is not None]
    return len(top_row)
