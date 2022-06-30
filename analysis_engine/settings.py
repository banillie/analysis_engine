from pathlib import Path
import platform

from datamaps.api import project_data_from_master, project_data_from_master_month


def _platform_docs_dir(dir: str) -> Path:
    #  Cross plaform file path handling. The dir (directorary) controls the report type.
    if platform.system() == "Linux":
        return Path.home() / "Documents" / dir
    if platform.system() == "Darwin":
        return Path.home() / "Documents" / dir
    else:
        return Path.home() / "Documents" / dir

#
# INITIATE_DICT = {
#     'cdg': {
#         'report': 'cdg',
#         'root_path': str(_platform_docs_dir('cdg')),
#         'config': '/core_data/cdg_config.ini',
#         'callable': project_data_from_master
#     },
#     'ipdc': {
#         'config': '/core_data/ipdc_config.ini',
#         'callable': project_data_from_master,
#     },
#     'top_250': {
#         'config': '/core_data/top_250_config.ini',
#         'callable': project_data_from_master_month,
#     }
# }  # controls the documents pointed to for reporting process via cli positional arguments.
#


def report_config(report_type: str):
    if report_type == 'cdg' or report_type == 'ipdc':
        func = project_data_from_master
    if report_type == 'top_250':
        func = project_data_from_master_month
    return {
        'report': report_type,
        'root_path': str(_platform_docs_dir(report_type)),
        'config': f'/core_data/{report_type}_config.ini',
        'callable': func,
        "master_path": "/core_data/json/master.json",
        "dashboard": "/input/dashboard_master.xlsx",
        "narrative_dashboard": "/input/narrative_dashboard_master.xlsx",
        "excel_save_path": "/output/{}.xlsx",
        "word_save_path": "/output/{}.docx",
        "word_landscape": "/input/summary_temp_landscape.docx",
    }
    # return INITIATE_DICT[report_type]


def set_default_args(op_args, port_group, default_quarter):
    if "group" not in op_args and "stage" not in op_args:
        op_args["group"] = port_group
    if "quarter" not in op_args:
        op_args["quarter"] = [default_quarter]

    return op_args



