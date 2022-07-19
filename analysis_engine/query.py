import datetime

from openpyxl import Workbook

from analysis_engine.milestones import MilestoneData, get_milestone_date

from analysis_engine.segmentation import get_iter_list, get_group, get_correct_p_data
from analysis_engine.render_utils import make_file_friendly
from analysis_engine.colouring import AMBER_FILL, SALMON_FILL


def data_query_into_wb(md, **kwargs) -> Workbook:
    """
    Returns data values for keys of interest. Keys placed on one page.
    Quarter data placed across different wbs.
    """

    wb = Workbook()
    iter_list = get_iter_list(md, **kwargs)
    for z, tp in enumerate(iter_list):
        # i = master.quarter_list.index(tp)  # handling here. for wrong quarter string
        ws = wb.create_sheet(
            make_file_friendly(tp)
        )  # creating worksheets. names restricted to 30 characters.
        ws.title = make_file_friendly(tp)  # title of worksheet
        group = get_group(md, tp, **kwargs)

        """list project names, groups and stage in ws"""
        for y, p in enumerate(group):  # p is project name
            p_data = get_correct_p_data(md, p, tp)
            abb = md['project_information'][p]['Abbreviations']
            ws.cell(row=2 + y, column=1).value = md['project_information'][p]["Directorate"]
            ws.cell(row=2 + y, column=2).value = p
            ws.cell(row=2 + y, column=3).value = abb
            ws.cell(row=2 + y, column=4).value = md['project_information'][p]["GMPP"]
            try:
                p_data_last = get_correct_p_data(
                    md, p, iter_list[z + 1]
                )
            except IndexError:
                p_data_last = None
            for x, key in enumerate(kwargs["key"]):
                ws.cell(row=1, column=5 + x, value=key)
                try:  # standard keys
                    value = p_data[key]
                    # value = convert_date(p_data[key])
                    if value is None:
                        ws.cell(row=2 + y, column=5 + x).value = "md"
                        ws.cell(row=2 + y, column=5 + x).fill = AMBER_FILL
                    else:
                        ws.cell(row=2 + y, column=5 + x, value=value)
                        if isinstance(value, datetime.datetime):
                            ws.cell(
                                row=2 + y, column=5 + x, value=value
                            ).number_format = "dd/mm/yy"

                    try:  # checks for change against next master in loop
                        lst_value = p_data_last[key]
                        # lst_value = convert_date(p_data_last[key])
                        if value != lst_value:
                            ws.cell(row=2 + y, column=5 + x).fill = SALMON_FILL
                    except (KeyError, UnboundLocalError, TypeError):
                        # KeyError is key not present in master.
                        # UnboundLocalError if there is no last_value.
                        # TypeError if project not in master. p_data_last becomes None.
                        pass
                except KeyError:  # milestone keys
                    if "quarter" in kwargs:
                        milestones_one = MilestoneData(md, quarter=[tp], group=[p])
                        try:
                            milestones_two = MilestoneData(
                                md, quarter=[iter_list[z + 1]], group=[p]
                            )
                        except IndexError:
                            pass
                    if "baseline" in kwargs:
                        milestones_one = MilestoneData(md, baseline=[tp], group=[p])
                        try:
                            milestones_two = MilestoneData(
                                md, baseline=[iter_list[z + 1]], group=[p]
                            )
                        except IndexError:
                            pass
                    date = get_milestone_date(
                        abb, milestones_one.milestone_dict, tp, key
                    )
                    if date is None:
                        ws.cell(row=2 + y, column=5 + x).value = "md"
                        ws.cell(row=2 + y, column=5 + x).fill = AMBER_FILL
                    else:
                        ws.cell(row=2 + y, column=5 + x).value = date
                        ws.cell(row=2 + y, column=5 + x).number_format = "dd/mm/yy"
                    try:  # checks for changes against next master in loop
                        lst_date = get_milestone_date(
                            abb,
                            milestones_two.milestone_dict,
                            iter_list[z + 1],
                            key,
                        )
                        if date != lst_date:
                            ws.cell(row=2 + y, column=5 + x).fill = SALMON_FILL
                    except (KeyError, UnboundLocalError, TypeError, IndexError):
                        pass

        ws.cell(row=1, column=1).value = "Group"
        ws.cell(row=1, column=2).value = "Project Name"
        ws.cell(row=1, column=3).value = "Project Acronym"
        ws.cell(row=1, column=4).value = "GMPP"

    wb.remove(wb["Sheet"])
    return wb


# def convert_date(date_str: str):
#     """
#     When date converted into json file the dates take the standard python format
#     year-month-day. This function converts format to year-day-month. This function is
#     used when the MilestoneData class is created. Seems to be the best place to deploy.
#     """
#     try:
#         return parser.parse(date_str)  # returns datetime
#     except TypeError:  # for a different data value e.g integer.
#         return date_str
#     except ValueError:  # for string data that is not a date.
#         return date_str
#     # is a ParserError necessary here also?