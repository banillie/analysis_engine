import os
import pytest
from datamaps.api import project_data_from_master
from openpyxl import load_workbook
from analysis_engine.data import open_word_doc, open_pickle_file


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
        os.path.join(
            os.getcwd(), "resources/" "test_project_info.xlsx"
        ),
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
