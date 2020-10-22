import os

import pytest
from datamaps.api import project_data_from_master

# from data_mgmt.data import open_word_doc
from database.database import create_db


@pytest.fixture
def db():
    db_path = os.path.join(os.getcwd(), "db_test.db")
    create_db(db_path)
    yield db_path
    os.remove(db_path)  # delete db


@pytest.fixture
def word_doc():
    wd_path = os.path.join(os.getcwd(), "resources/summary_temp.docx")
    doc = open_word_doc(wd_path)
    return doc


@pytest.fixture
def contact_master():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                               "contact_info_master_3_2019.xlsx"), 3, 2019),
            project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                               "contact_info_master_4_2019.xlsx"), 4, 2019)]


@pytest.fixture
def spent_master():
    return project_data_from_master(os.path.join(os.getcwd(), "resources/spent_data_master_2_2020.xlsx"), 2, 2020)


@pytest.fixture
def one_master():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                               "milestones_test_master_4_2019.xlsx"), 4, 2019)]


@pytest.fixture
def master_path_apostrophe():
    return os.path.join(os.getcwd(), "resources/"
                                     "one_row_master.xlsx")


@pytest.fixture
def project_info():
    return project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                              "test_project_group_id_no.xlsx"), 1, 2099)


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
def benefits_master():
    master = [
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "benefits_master_test_2_2020.xlsx"), 2, 2020),
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "benefits_master_test_1_2020.xlsx"), 1, 2020)
    ]
    return master


@pytest.fixture()
def project_group_id():
    return project_data_from_master(os.path.join(os.getcwd(), "resources/test_project_group_id_no.xlsx"), 1, 2099)


@pytest.fixture()
def dca_masters():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2019_dcas.xlsx"), 1, 2099),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2018_dcas.xlsx"), 1, 2098),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2017_dcas.xlsx"), 1, 2097),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2016_dcas.xlsx"), 1, 2096)]


@pytest.fixture()
def costs_masters():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/cost_test_master_4_2019.xlsx"), 4, 2019),
            project_data_from_master(os.path.join(os.getcwd(), "resources/cost_test_master_4_2018.xlsx"), 4, 2018)]
