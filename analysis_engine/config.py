from datamaps.api import project_data_from_master, project_data_from_master_month


INITIATE_DICT = {
    'cdg': {
        'report': 'cdg',
        'config': '/core_data/cdg_config.ini',
        'callable': project_data_from_master
    },
    'ipdc': {
        'config': '/core_data/ipdc_config.ini',
        'callable': project_data_from_master,
    },
    'top_250': {
        'config': '/core_data/top_250_config.ini',
        'callable': project_data_from_master_month,
    }
}  # controls the documents pointed to for reporting process via cli positional arguments.


def report_config(report_type: str):
    return INITIATE_DICT[report_type]




