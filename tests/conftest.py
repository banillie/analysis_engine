import os

import pytest

from vfm.database import create_db


@pytest.fixture
def db():
    db_path = os.path.join(os.getcwd(), "db_test.db")
    create_db(db_path)
    yield db_path
    os.remove(db_path)  # delete db


@pytest.fixture
def master_path():
    return os.path.join(os.getcwd(), "resources/" 
           "cut_down_master_4_2016.xlsx")


