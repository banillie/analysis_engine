import sqlite3


class MilestoneData:
    def __init__(self, db):
        self.db = db

    def get_project_milestones(self, project_name):
        conn = sqlite3.connect(self.db)
        c = conn.cursor()

        return c.execute(f"SELECT * FROM milestone WHERE project_name='{project_name}'").fetchall()



