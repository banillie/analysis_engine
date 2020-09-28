import sqlite3
from vfm.database import import_master_to_db




def get_cursor(db, master_path):
    import_master_to_db(db, master_path)
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


def test_import_master_to_db(db, master_path):
    c = get_cursor(db, master_path)
    c.execute("""SELECT count(*) FROM project""")
    assert c.fetchall() == [(6,)]


def test_apostrophe_in_text(db, master_path):
    c = get_cursor(db, master_path)
    c.execute("""SELECT notes FROM milestone WHERE project_name = 'Apollo 11'""")
    assert ("Don't you know an apparition is just a cheap date. " \
           "What have you been drinking these days") in c.fetchall()[0][0]


def test_insert_quarter_data_with_foreign_keys(db, master_path):
    c = get_cursor(db, master_path)
    c.execute("""SELECT quarter_id, group_id FROM project WHERE name = 'Apollo 11'""")
    assert c.fetchall() == [('Q4 19/20', 'HSMRPG')]


def test_insert_milestone_data_with_foreign_keys(db, master_path):
    c = get_cursor(db, master_path)
    c.execute("""SELECT milestone_type_id, quarter_id, project_id FROM milestone WHERE project_name = 'Apollo 11' 
    and milestone_type_id = 'Approval'""")
    assert c.fetchall() == [('Approval', 'Q4 19/20', 2)]
    c.execute("""SELECT milestone_type_id, quarter_id, project_id FROM milestone WHERE project_name = 'Apollo 11' 
        and milestone_type_id = 'Assurance'""")
    assert c.fetchall() == [('Assurance', 'Q4 19/20', 2)]


def test_sqlite_select_commands_across_tables(db, master_path):
    c = get_cursor(db, master_path)
    c.execute(
        """select milestone.name from milestone, project where 
        milestone.project_name = project.name and project.group_id = 'AMIS'""")
    assert c.fetchall() == [('Earth Command',), ('Inverted Cosmonauts',), ('Sputnik Sea',), ('Team Magma',)]