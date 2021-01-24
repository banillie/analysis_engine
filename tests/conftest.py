import os

import pytest
from datamaps.api import project_data_from_master

from analysis_engine.data import open_word_doc
from other.database.database import create_db


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
def one_milestones_master():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                               "milestones_test_master_4_2019.xlsx"), 4, 2019)]


@pytest.fixture
def two_masters():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_1_2020_f9.xlsx"), 1, 2020),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2019_f9.xlsx"), 4, 2019)]


@pytest.fixture
def master_path_apostrophe():
    return os.path.join(os.getcwd(), "resources/"
                                     "one_row_master.xlsx")


@pytest.fixture
def project_info():
    return project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                              "test_project_group_id_no.xlsx"), 1, 2099)


@pytest.fixture
def project_info_incorrect():
    return project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                              "test_project_group_id_no_incorrect.xlsx"), 1, 2099)


@pytest.fixture()
def basic_masters_dicts():
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
def basic_pickle():
    return os.path.join(os.getcwd(), "resources/test_master.pickle")


@pytest.fixture()
def basic_masters_file_paths():
    file_paths = [
        os.path.join(os.getcwd(), "resources/"
                                  "cut_down_master_4_2016.xlsx"),
        os.path.join(os.getcwd(), "resources/"
                                  "cut_down_master_4_2017.xlsx"),
        os.path.join(os.getcwd(), "resources/"
                                  "cut_down_master_4_2018.xlsx"),
    ]
    return file_paths


@pytest.fixture()
def basic_master_wrong_baselines():
    test_master_data = [
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "cut_down_master_4_2017_incorrect_baselines.xlsx"), 4, 2017),
        project_data_from_master(os.path.join(os.getcwd(), "resources/"
                                                           "cut_down_master_4_2018_incorrect_baselines.xlsx"), 4, 2018)

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
def project_group_id_path():
    return os.path.join(os.getcwd(), "resources/test_project_group_id_no.xlsx")


@pytest.fixture()
def project_old_fy_path():
    return os.path.join(os.getcwd(), "resources/test_project_old_fy_data.xlsx")


@pytest.fixture()
def dca_masters():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2019_dcas.xlsx"), 4, 2019),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2018_dcas.xlsx"), 4, 2018),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2017_dcas.xlsx"), 4, 2017),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_master_4_2016_dcas.xlsx"), 4, 2016)]


@pytest.fixture()
def risk_masters():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/test_risk_master_2_2020.xlsx"), 2, 2020),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_risk_master_1_2020.xlsx"), 1, 2020)]


@pytest.fixture()
def costs_masters():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/cost_test_master_1_2020.xlsx"), 1, 2020),
            project_data_from_master(os.path.join(os.getcwd(), "resources/cost_test_master_4_2019.xlsx"), 4, 2019)]


@pytest.fixture()
def vfm_masters():
    return [project_data_from_master(os.path.join(os.getcwd(), "resources/test_vfm_master_1_2020.xlsx"), 1, 2020),
            project_data_from_master(os.path.join(os.getcwd(), "resources/test_vfm_master_4_2019.xlsx"), 4, 2019)]


@pytest.fixture()
def change_log():
    return os.path.join(os.getcwd(), "resources/test_key_change_log.xlsx")


@pytest.fixture()
def list_cost_masters_files():
    return [
        os.path.join(os.getcwd(), "resources/cost_test_master_4_2018.xlsx")
    ]


@pytest.fixture()
def list_test_masters_files():
    return [
        os.path.join(os.getcwd(), "resources/test_master_1_2020.xlsx"),
        os.path.join(os.getcwd(), "resources/test_master_4_2019.xlsx"),
        os.path.join(os.getcwd(), "resources/test_master_4_2018.xlsx"),
        os.path.join(os.getcwd(), "resources/test_master_4_2017.xlsx"),
        os.path.join(os.getcwd(), "resources/test_master_4_2016.xlsx")
    ]
