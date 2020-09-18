from vfm.database import create_db
import os

#  test that dB is created.
def test_create_db():
    create_db('db_test')
    import sqlite3
    CWD_PATH = os.getcwd()
    db_path = os.path.join(CWD_PATH, "db_test.db")
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("INSERT INTO quarter VALUES (1, 'test_quarter', 1)")
    conn.commit()
    c.execute("""
        SELECT count(*) FROM 'quarter'
        """)
    assert c.fetchall() == [(1,)]
    os.remove(db_path) #  delete db
