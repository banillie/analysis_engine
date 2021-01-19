import sqlite3
import pytest
from other.database.database import import_master_to_db
from other.database.shedbase import MilestoneData



def test_create_db(db):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    c.execute("INSERT INTO quarter (quarter_id, quarter_number) VALUES ('test_quarter', 1)")
    conn.commit()
    c.execute("""
        SELECT count(*) FROM 'quarter'
        """)
    assert c.fetchall() == [(1,)]


def test_import_master_to_db(db, one_milestones_master, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, one_milestones_master, project_group_id)
    c.execute("""SELECT count(*) FROM project""")
    assert c.fetchall() == [(6,)]


def test_explore_list_of_single_values_into_dft_group(db, milestone_masters, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, milestone_masters, project_group_id)
    c.execute("""SELECT * FROM dft_group""")
    assert c.fetchall() == [(1, 'Rail Group'), (2, 'HSMRPG'), (3, 'RPE'), (4, 'AMIS')]


def test_apostrophe_in_text(db, milestone_masters, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, milestone_masters, project_group_id)
    c.execute("""SELECT notes FROM milestone WHERE project_name = 'Apollo 11'""")
    assert ("Don't you know an apparition is just a cheap date. " \
           "What have you been drinking these days") in c.fetchall()[0][0]


def test_explore_quarter_data_with_foreign_keys(db, milestone_masters, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, milestone_masters, project_group_id)
    c.execute("""SELECT group_name FROM project WHERE name = 'Apollo 11'""")
    assert c.fetchall() == [('HSMRPG',)]


def test_explore_milestone_data_with_foreign_keys(db, milestone_masters, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, milestone_masters, project_group_id)
    c.execute("""SELECT milestone_type, quarter_id, project_id FROM milestone WHERE project_name = 'Apollo 11' 
    and milestone_type = 'Approval'""")
    assert c.fetchall() == [('Approval', 'Q4 19/20', 'B2'), ('Approval', 'Q4 18/19', 'B2')]
    c.execute("""SELECT milestone_type, quarter_id, project_id FROM milestone WHERE project_name = 'Apollo 11' 
        and milestone_type = 'Assurance'""")
    assert c.fetchall() == [('Assurance', 'Q4 19/20', 'B2'), ('Assurance', 'Q4 18/19', 'B2')]


def test_explore_sqlite_select_commands_across_tables(db, milestone_masters, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, milestone_masters, project_group_id)
    c.execute(
        """select milestone.name from milestone, project where 
        milestone.project_name = project.name and project.group_name = 'Rail Group'
        and milestone.milestone_type = 'Approval' and milestone.quarter_id = 'Q4 19/20'""")
    assert c.fetchall() == [('Mercury Lade',), ('Sputnik Sea',)]
    c.execute(
        """select milestone.name from milestone, project where 
        milestone.project_name = project.name and project.group_name = 'AMIS'
        and milestone.milestone_type = 'Project' and milestone.quarter_id = 'Q4 19/20'""")
    assert c.fetchall() == [('Orbital Landing',)]

def test_explore_more_than_one_quarter_master_project_names(db, milestone_masters, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, milestone_masters, project_group_id)
    c.execute("""SELECT * FROM project WHERE name = 'Mars'""")
    assert c.fetchall() == [(1, 'AMIS', 'D6', 'Mars'),]

def test_insert_into_project_table_throws_unqiue_acception(db, milestone_masters, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, milestone_masters, project_group_id)
    with pytest.raises(sqlite3.IntegrityError):
        c.execute("INSERT INTO project (group_name, project_id, name) "
                  "VALUES ('Rail Group', 'D6', 'Mars')")
        c.execute("INSERT INTO project (group_name, project_id, name) "
                  "VALUES ('Rail Group', 'D6', 'Jupiter')")
        c.execute("INSERT INTO project (group_name, project_id, name) "
                  "VALUES ('Rail Group', 'E6', 'Mars')")

def test_get_project_milestones(db, milestone_masters, project_group_id):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    import_master_to_db(db, milestone_masters, project_group_id)
    ms = MilestoneData(db)
    ms_mars = ms.get_project_milestones('Mars')
    assert ms_mars == [(1, 'Approval', 'Q4 19/20', 'D6', 'Mars', 'Meteorite Shuttle', 'wood', 1.0, '2020-03-10',
                             '2020-03-10', 0.0, 'Complete', 'What you see if all there is', 'None', 'None'), (
                            2, 'Assurance', 'Q4 19/20', 'D6', 'Mars', 'Tranquility Hypatia', 'None', 'None',
                            '2020-09-21', '2020-09-21', 0.0, 'Live', 'What you see if all there is', 'wood', 'None'), (
                            3, 'Project', 'Q4 19/20', 'D6', 'Mars', 'Orbital Landing', 'None', 'None', '2019-12-16',
                            '2019-12-16', 0.0, 'Complete', 'The sea gets deeper the further you go into it', 'None',
                            'Yes')]



