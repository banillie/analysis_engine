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
    c.execute("INSERT INTO quarter VALUES ('test_quarter', 1)")
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
