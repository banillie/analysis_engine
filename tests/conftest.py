import os

import pytest

from database.database import create_db

from datamaps.api import project_data_from_master


@pytest.fixture
def db():
    db_path = os.path.join(os.getcwd(), "db_test.db")
    create_db(db_path)
    yield db_path
    os.remove(db_path)  # delete db


@pytest.fixture
def one_master():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                               "milestones_test_master_4_2019.xlsx"), 4, 2019)]


@pytest.fixture
def master_path_apostrophe():
    return os.path.join(os.getcwd(), "resources/"
                                     "one_row_master.xlsx")


@pytest.fixture
def abbreviations():
    return {'Sea of Tranquility': 'SoT',
            'Apollo 11': 'A11',
            'Apollo 13': 'A13',
            'Falcon 9': 'F9',
            'Columbia': 'Columbia',
            'Mars': 'Mars'}


@pytest.fixture()
def basic_master():
    test_master_data = [
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "cut_down_master_4_2016.xlsx"), 4, 2016),
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "cut_down_master_4_2017.xlsx"), 4, 2017),
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "cut_down_master_4_2018.xlsx"), 4, 2018)

    ]
    return test_master_data


@pytest.fixture()
def milestone_masters():
    test_master_data = [
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "milestones_test_master_4_2019.xlsx"), 4, 2019),
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "milestones_test_master_4_2018.xlsx"), 4, 2018)
    ]
    return test_master_data


@pytest.fixture()
def diff_milestone_types():
    master = [
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "diff_milestone_data_formats_master_2_2020.xlsx"), 2, 2020),
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "diff_milestone_data_formats_master_1_2020.xlsx"), 1, 2020)
    ]
    return master


@pytest.fixture()
def project_group_id():
    return project_data_from_master(os.path.join(os.getcwd(), "resources/test_project_group_id_no.xlsx"), 1, 2099)
