import sqlite3
from vfm.database import import_master_to_db


#  test that dB is created.
def test_create_db(db):
    conn = sqlite3.connect(db)
    c = conn.cursor()
    c.execute("INSERT INTO quarter VALUES (1, 'test_quarter', 1)")
    conn.commit()
    c.execute("""
        SELECT count(*) FROM 'quarter'
        """)
    assert c.fetchall() == [(1,)]


def test_import_master_to_db(db, master_path):
    import_master_to_db(db, master_path)
    conn = sqlite3.connect(db)
    c = conn.cursor()
    c.execute("""SELECT count(*) FROM project""")
    assert c.fetchall() == [(5,)]



