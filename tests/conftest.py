import os

import pytest

from vfm.database import create_db

from datamaps.api import project_data_from_master


@pytest.fixture
def db():
    db_path = os.path.join(os.getcwd(), "db_test.db")
    create_db(db_path)
    yield db_path
    os.remove(db_path)  # delete db


@pytest.fixture
def master_path():
    return os.path.join(os.getcwd(), "resources/" 
           "milestones_test_master_4_2019.xlsx")


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

# group of masters
@pytest.fixture()
def mst():
    test_master_data = [
            project_data_from_master(os.path.join(os.getcwd(), "resources/"
        "cut_down_master_4_2016.xlsx"), 4, 2016),
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
        "cut_down_master_4_2017.xlsx"), 4, 2017)
        ]
    return test_master_data