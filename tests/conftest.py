import os
import pytest
from datamaps.api import project_data_from_master
from openpyxl import load_workbook

from analysis_engine.cdg_data import (
    cdg_root_path,
    cdg_get_master_data,
    cdg_get_project_information,
)
from analysis_engine.data import open_word_doc, open_pickle_file, root_path
from analysis_engine.top35_data import top35_root_path, top35_get_master_data, top35_get_project_information


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
    return project_data_from_master(
        os.path.join(os.getcwd(), "resources/" "test_project_info.xlsx"),
        1,
        2020,
    )


def cdg_project_info():
    return project_data_from_master(
        os.path.join(os.getcwd(), "resources/" "test_cdg_proj_info.xlsx"),
        1,
        2020,
    )


@pytest.fixture()
def basic_masters_dicts():
    test_master_data = [
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "cut_down_master_4_2016.xlsx"),
            4,
            2016,
        ),
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "cut_down_master_4_2017.xlsx"),
            4,
            2017,
        ),
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "cut_down_master_4_2018.xlsx"),
            4,
            2018,
        ),
    ]
    return test_master_data


@pytest.fixture()
def full_test_masters_dict():
    test_master_data = [
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "test_master_1_2020.xlsx"), 1, 2020
        ),
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "test_master_4_2019.xlsx"), 4, 2019
        ),
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "test_master_4_2018.xlsx"), 4, 2018
        ),
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "test_master_4_2017.xlsx"), 4, 2017
        ),
        project_data_from_master(
            os.path.join(os.getcwd(), "resources/" "test_master_4_2016.xlsx"), 4, 2016
        ),
    ]
    return test_master_data


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
def json_path():
    return os.path.join(os.getcwd(), "resources/" "json_master")


@pytest.fixture()
def master_pickle():
    return open_pickle_file(os.path.join(os.getcwd(), "resources/test_master.pickle"))


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


@pytest.fixture()
def cdg_data():
    return {
        "test": {
            "docx_save_path": "resources/{}.docx",
            "data": (cdg_masters(), cdg_project_info()),
            "op_args": {
                "quarter": ["Q4 20/21"],
                "group": ["CFPD"],
                "chart": True,
                "data_type": "cdg",
            },
        },
    }


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


def top35_master():
    master_data_list = [
        project_data_from_master("resources/250_master_test_2.xlsx", 4, 2020
        ),
        project_data_from_master("resources/250_master_test_1.xlsx", 3, 2020
        ),
    ]
    return master_data_list


def top35_project_information():
    return project_data_from_master("resources/250_project_info_test.xlsx", 2, 2020)


@pytest.fixture()
def top35_data():
    return {
        "docx_save_path": "resources/{}.docx",
            "data": (top35_master(), top35_project_information()),
            "op_args": {
                "quarter": ["Q4 20/21"],
                # "group": ["HSRG", "RSS", "RIG", "RPE"],
                "group": ["RIG"],
                "chart": True,
                "data_type": "top35",
                "circle_colour": 'No',
            },
        }