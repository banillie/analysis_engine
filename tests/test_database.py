import sqlite3
from database.database import import_master_to_db

#  refactor. don't need milestone_master as an argument
def get_cursor(db, milestone_masters):
    import_master_to_db(db, milestone_masters)
    conn = sqlite3.connect(db)
    c = conn.cursor()
    return c


def test_create_db(db):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    c.execute("INSERT INTO quarter (quarter_id, quarter_number) VALUES ('test_quarter', 1)")
    conn.commit()
    c.execute("""
        SELECT count(*) FROM 'quarter'
        """)
    assert c.fetchall() == [(1,)]


def test_import_master_to_db(db, one_master):
    c = get_cursor(db, one_master)
    c.execute("""SELECT count(*) FROM project""")
    assert c.fetchall() == [(6,)]


def test_insert_list_of_single_values_into_dft_group(db, milestone_masters):
    c = get_cursor(db, milestone_masters)
    c.execute("""SELECT * FROM dft_group""")
    assert c.fetchall() == [(1, 'Rail Group'), (2, 'HSMRPG'), (3, 'RPE'), (4, 'AMIS')]


def test_apostrophe_in_text(db, milestone_masters):
    c = get_cursor(db, milestone_masters)
    c.execute("""SELECT notes FROM milestone WHERE project_name = 'Apollo 11'""")
    assert ("Don't you know an apparition is just a cheap date. " \
           "What have you been drinking these days") in c.fetchall()[0][0]


def test_insert_quarter_data_with_foreign_keys(db, milestone_masters):
    c = get_cursor(db, milestone_masters)
    c.execute("""SELECT group_name FROM project WHERE name = 'Apollo 11'""")
    assert c.fetchall() == [('HSMRPG',)]


def test_insert_milestone_data_with_foreign_keys(db, milestone_masters):
    c = get_cursor(db, milestone_masters)
    c.execute("""SELECT milestone_type, quarter_id, project_id FROM milestone WHERE project_name = 'Apollo 11' 
    and milestone_type = 'Approval'""")
    assert c.fetchall() == [('Approval', 'Q4 19/20', 'B2'), ('Approval', 'Q4 18/19', 'B2')]
    c.execute("""SELECT milestone_type, quarter_id, project_id FROM milestone WHERE project_name = 'Apollo 11' 
        and milestone_type = 'Assurance'""")
    assert c.fetchall() == [('Assurance', 'Q4 19/20', 'B2'), ('Assurance', 'Q4 18/19', 'B2')]


def test_sqlite_select_commands_across_tables(db, milestone_masters):
    c = get_cursor(db, milestone_masters)
    c.execute(
        """select milestone.name from milestone, project where 
        milestone.project_name = project.name and project.group_name = 'AMIS'
        and milestone.milestone_type = 'Approval' and milestone.quarter_id = 'Q4 19/20'""")
    assert c.fetchall() == [('Earth Command',), ('Sputnik Sea',)]
    c.execute(
        """select milestone.name from milestone, project where 
        milestone.project_name = project.name and project.group_name = 'AMIS'
        and milestone.milestone_type = 'Assurance' and milestone.quarter_id = 'Q4 19/20'""")
    assert c.fetchall() == [('Inverted Cosmonauts',), ('Team Magma',)]
    c.execute(
        """select milestone.name from milestone, project where 
        milestone.project_name = project.name and project.group_name = 'AMIS'
        and milestone.milestone_type = 'Project' and milestone.quarter_id = 'Q4 19/20'""")
    assert c.fetchall() == [('Standard A',), ('Standard A',)]

def test_more_than_one_quarter_master_project_names(db, milestone_masters):
    c = get_cursor(db, milestone_masters)
    c.execute("""SELECT * FROM project WHERE name = 'Mars'""")
    assert c.fetchall() == [(1, 'Rail Group', 'D6', 'Mars'),]