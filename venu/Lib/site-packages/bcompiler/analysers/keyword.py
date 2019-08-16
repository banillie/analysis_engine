"""
Analyser for outputting ad hoc data from a master, based on given keyword.
"""
import reprlib
import sys

from openpyxl import load_workbook, Workbook

from .utils import MASTER_XLSX, logger, get_number_of_projects
from ..utils import runtime_config, CONFIG_FILE, project_data_from_master

runtime_config.read(CONFIG_FILE)


def process_master(source_wb, project_number, search_term) -> list:
    source_sheet = source_wb.active
    project_name = source_sheet.cell(row=1, column=project_number).value
    p_data = project_data_from_master(source_wb, opened_wb=True)[project_name]
    data = [item for item in p_data.items() if search_term in item[0]]
    if not data:
        logger.warning(f"No matching keyword found in {project_name}")
    return (project_name, data)


def run(output_path=None, user_provided_master_path=None, search_term: str=None, xlsx: bool=False):
    if user_provided_master_path:
        logger.info(f"Using master file: {user_provided_master_path}")
        wb = load_workbook(user_provided_master_path)
    else:
        logger.info(f"Using default master file (refer to config.ini)")
        wb = load_workbook(MASTER_XLSX)

    project_count = get_number_of_projects(wb)

    if not xlsx:
        print("{:<50}{:<50}{:<10}".format("Project", "Key", "Value"))
        print("{:*<140}".format(""))

        r = reprlib.Repr()
        r.maxstring = 48

        for p in range(2, project_count + 2):

            # do the work
            project_name, data = process_master(wb, p, search_term)

            for item in data:
                print("{:<50}{:<50}{:<10}".format(
                    r.repr(project_name),
                    r.repr(item[0]),
                    r.repr(item[1])))
    else:
        output_wb = Workbook()
        ws = output_wb.active
        ws.title = "Results of search"

        start_row = 1

        def val_gen(row, p_name, t, ws):
            yield ws.cell(column=1, row=row, value=p_name)
            yield ws.cell(column=2, row=row, value=t[0])
            yield ws.cell(column=3, row=row, value=t[1])

        for p in range(2, project_count + 2):
            # do the work
            project_name, data = process_master(wb, p, search_term)
            logger.info(f"Processing {project_name}")

            for i, item in enumerate(data, start_row):
                g = val_gen(start_row, project_name, item, ws)
                for cell in range(1, 4):
                    next(g)
                start_row += 1
        output_wb.save(xlsx[0])


if __name__ == '__main__':
    run()
