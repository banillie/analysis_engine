"""
Analyser for outputting ad hoc data from a master.
"""
import reprlib

from openpyxl import load_workbook

from .utils import MASTER_XLSX, logger, get_number_of_projects
from ..utils import runtime_config, CONFIG_FILE, project_data_from_master

runtime_config.read(CONFIG_FILE)


def process_master(source_wb, project_number, search_term) -> list:
    source_sheet = source_wb.active
    project_name = source_sheet.cell(row=1, column=project_number).value
    p_data = project_data_from_master(source_wb, opened_wb=True)[project_name]
    data = [item for item in p_data.items() if search_term in item[0]]
    return (project_name, data)


def run(output_path=None, user_provided_master_path=None, search_term: str=None):
    if user_provided_master_path:
        logger.info(f"Using master file: {user_provided_master_path}")
        wb = load_workbook(user_provided_master_path)
    else:
        logger.info(f"Using default master file (refer to config.ini)")
        wb = load_workbook(MASTER_XLSX)

    project_count = get_number_of_projects(wb)

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


if __name__ == '__main__':
    run(search_term='working')
