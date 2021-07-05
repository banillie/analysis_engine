import configparser
import os
import pytest
from datamaps.api import project_data_from_master, project_data_from_master_month
from openpyxl import load_workbook

from analysis_engine.data import open_word_doc, open_pickle_file, root_path, Master, open_json_file, JsonMaster, \
    get_master_data, get_project_information, JsonData


def pytest_addoption(parser):
    parser.addoption(
        "--runslow", action="store_true", default=False, help="run slow tests"
    )


def pytest_configure(config):
    config.addinivalue_line("markers", "slow: mark test as slow to run")


def pytest_collection_modifyitems(config, items):
    if config.getoption("--runslow"):
        # --runslow given in cli: do not skip slow tests
        return
    skip_slow = pytest.mark.skip(reason="need --runslow option to run")
    for item in items:
        if "slow" in item.keywords:
            item.add_marker(skip_slow)


@pytest.fixture
def word_doc():
    wd_path = os.path.join(os.getcwd(), "resources/summary_temp.docx")
    doc = open_word_doc(wd_path)
    return doc


@pytest.fixture
def word_doc_landscape():
    wd_path = os.path.join(os.getcwd(), "resources/summary_temp_landscape.docx")
    doc = open_word_doc(wd_path)
    return doc


@pytest.fixture
def project_info():
    return get_project_information(
        "resources/basic_m_confi.ini",
        "resources/"
    )


# def cdg_project_info():
#     return project_data_from_master(
#         os.path.join(os.getcwd(), "resources/" "test_cdg_proj_info.xlsx"),
#         1,
#         2020,
#     )


@pytest.fixture()
def basic_masters_dicts():
    return get_master_data(
        "resources/basic_m_confi.ini",
        "resources/",
        project_data_from_master
    )


@pytest.fixture()
def full_test_masters_dict():
    return get_master_data(
        "resources/full_m_confi.ini",
        "resources/",
        project_data_from_master
    )


@pytest.fixture()
def one_master_dict():
    return project_data_from_master(
        os.path.join(os.getcwd(), "resources/" "test_master_1_2020.xlsx"), 1, 2020
    )


def cdg_masters():
    data = [
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "test_cdg_master_Q4.xlsx"), 4, 2020
        ),
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "test_cdg_master_Q3.xlsx"), 3, 2020
        ),
    ]
    return data


@pytest.fixture()
def master_json_path():
    return os.path.join(os.getcwd(), "resources/" "json_master")


@pytest.fixture()
def top35_master_json_path():
    return os.path.join(os.getcwd(), "resources/" "top250_json_master")


@pytest.fixture()
def master():
    jm = open_json_file(os.path.join(os.getcwd(), "resources/" "json_master.json"))
    return Master(jm)


@pytest.fixture()
def master_pickle_file_path():
    return os.path.join(os.getcwd(), "resources/test_master.pickle")


@pytest.fixture()
def basic_masters_file_paths():
    file_paths = [
        os.path.join(os.getcwd(), "resources/" "cut_down_master_4_2016.xlsx"),
        os.path.join(os.getcwd(), "resources/" "cut_down_master_4_2017.xlsx"),
        os.path.join(os.getcwd(), "resources/" "cut_down_master_4_2018.xlsx"),
    ]
    return file_paths


@pytest.fixture()
def change_log():
    return os.path.join(os.getcwd(), "resources/test_key_change_log.xlsx")


@pytest.fixture()
def list_cost_masters_files():
    return [os.path.join(os.getcwd(), "resources/cost_test_master_4_2018.xlsx")]


@pytest.fixture()
def dashboard_template():
    return load_workbook(
        os.path.join(os.getcwd(), "resources/test_dashboards_master.xlsx")
    )


@pytest.fixture()
def key_file():
    return os.path.join(os.getcwd(), "resources/test_key_names.csv")


@pytest.fixture()
def horizontal_bar_chart_data():
    return os.path.join(os.getcwd(), "resources/horizontal_bar_chart_manual_data.xlsx")


@pytest.fixture()
def sp_data():
    return os.path.join(os.getcwd(), "resources/test_sp_master.xlsx")


# @pytest.fixture()
# def cdg_data():
#     return {
#         "test": {
#             "docx_save_path": "resources/{}.docx",
#             "data": (cdg_masters(), cdg_project_info()),
#             "op_args": {
#                 "quarter": ["Q4 20/21"],
#                 "group": ["CFPD"],
#                 "chart": True,
#                 "data_type": "cdg",
#             },
#         },
#     }


@pytest.fixture()
def ipdc_data():
    return {
        "docx_save_path": "resources/{}.docx",
        "master": open_pickle_file(os.path.join(os.getcwd(), "resources/test_master.pickle")),
        "op_args": {
            "quarter": ["Q1 20/21"],
            "group": ["HSRG", "RSS", "RIG", "AMIS", "RPE"],
            "chart": True,
        },
    }


@pytest.fixture()
def top35_master():
    return get_master_data(
            "resources/top250_confi.ini",
            "resources/",
            project_data_from_master_month
        )


@pytest.fixture()
def top35_project_info():
    return get_project_information(
            "resources/top250_confi.ini",
            "resources/"
        )


@pytest.fixture()
def top35_data():
    return {
        "docx_save_path": "resources/{}.docx",
        "master": Master(open_json_file("resources/top250_json_master.json")),
        "op_args": {
            # "quarter": ["Month(June), 2021"],
            "quarter": ["standard"],
            # "group": ["HSRG", "RSS", "RIG", "RPE"],
            "group": ["RIG"],
            "chart": True,
            "data_type": "top35",
            "circle_colour": 'No',
        },
    }

