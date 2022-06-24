from datamaps.api import project_data_from_master

from analysis_engine.core_data import _platform_docs_dir


REPORTING_TYPE = "cdg"

ROOT_PATH = _platform_docs_dir(REPORTING_TYPE)

INITIATE_DICT = {
    "cdg": {"config": "/core_data/cdg_config.ini", "callable": project_data_from_master}
}

CONFIG_PATH = str(ROOT_PATH) + INITIATE_DICT[REPORTING_TYPE]["config"]

OP_ARGS_DICT = {
    # "docx_save_path": str(cdg_root_path / "output/{}.docx"),
    # "master": Master(open_json_file(str(cdg_root_path / "core_data/json/master.json"))),
    # "op_args": {
        "quarter": ["Q4 21/22"],
        # "quarter": ["standard"],
        "group": ["SCS", "CFPD", "GF"],
        # "group": ["SCS", "GF"],
        "chart": "show",
        "data_type": "cdg",
        # "type": "income",
        "blue_line": "CDG",
        "dates": ["1/10/2021", "1/10/2022"],
        "fig_size": "half_horizontal",
        "rag_number": "5",
        # "order_by": "cost",
        "angles": [300, 360, 60],
        "none_handle": "none",
    # },
    #     "dashboard": str(cdg_root_path / "input/dashboard_master.xlsx"),
    #     "narrative_dashboard": str(cdg_root_path / "input/narrative_dashboard_master.xlsx"),
    #     "excel_save_path": str(cdg_root_path / "output/{}.xlsx"),
    #     "word_save_path": str(cdg_root_path / "output/{}.docx")
}
