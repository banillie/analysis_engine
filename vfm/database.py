import sqlite3


def create_connect_db(db_name):
    conn = sqlite3.connect(db_name + '.db')
    return conn


#  create a new table in vfm db.
def create_vfm_table(db_name, insert_quarter):
    conn = sqlite3.connect(db_name + '.db')
    c = conn.cursor()

    c.execute("""CREATE TABLE '{quarter}'
            (project_name text,
            project_group text,
            npv real,
            adjusted_bcr real,
            initial_bcr real,
            vfm_cat_single text,
            pvc real,
            pvb real)""".format(quarter=insert_quarter))

    conn.commit()
    conn.close()

#  create a new table in vfm db.
def create_db(db_name):
    conn = sqlite3.connect(db_name + '.db')
    c = conn.cursor()
    c.execute("""
    DROP TABLE IF EXISTS quarter;
    """)
    c.execute("""
    DROP TABLE IF EXISTS project_group;
    """)
    c.execute("""
    DROP TABLE IF EXISTS project;
    """)
    c.execute("""
    DROP TABLE IF EXISTS milestone_type;
    """)
    c.execute("""
    DROP TABLE IF EXISTS milestone;
    """)

    c.execute("""CREATE TABLE 'quarter'
            (id integer primary key, 
            name text,
            quarter_number integer)""")

    c.execute("""CREATE TABLE 'project_group'
                (id integer primary key,
                name text)""")

    c.execute("""CREATE TABLE 'project'
            (id integer primary key,
            quarter_id integer,
            group_id integer,
            project_name text,
            FOREIGN KEY(quarter_id) REFERENCES quarter(id)
            FOREIGN KEY(group_id) REFERENCES project_group(id))""")

    c.execute("""CREATE TABLE 'milestone_type'
            (id integer primary key,
            type text,
            description text)""")

    c.execute("""CREATE TABLE 'milestone'
            (id integer primary key,
            milestone_type_id text,
            project_id integer,
            gov_type text,
            ver_no real,
            orig_baseline text,
            forecast_actual text,
            variance real,
            status text,
            notes text,
            FOREIGN KEY(project_id) REFERENCES project(id)
            FOREIGN KEY(milestone_type_id) REFERENCES milestone_type(id)
            )""")

    conn.commit()
    conn.close()



#  gets vfm data values from master data in excel wbs.
def get_vfm_values(masters):
    output_list = []
    for master in masters:
        quarter = master.quarter
        for project in master.projects:
            group = master.data[project]['DfT Group']
            npv = master.data[project]['NPV for all projects ' \
                                       'and NPV for programmes if available']
            adjusted_bcr = master.data[project]['Adjusted Benefits Cost Ratio (BCR)']
            initial_bcr = master.data[project]['Initial Benefits Cost Ratio (BCR)']
            vfm_cat_single = master.data[project]['VfM Category single entry']
            pvc = master.data[project]['Present Value Cost (PVC)']
            pvb = master.data[project]['Present Value Benefit (PVB)']
            output_list.append((str(quarter),
                                project,
                                group,
                                npv,
                                adjusted_bcr,
                                initial_bcr,
                                vfm_cat_single,
                                pvc,
                                pvb))

    return output_list


#  returns vfm list of tuples for specified quarter
def get_quarter_values(vfm_list, quarter):
    output_list = []
    for i in vfm_list:
        if i[0] == quarter:
            output_list.append(i[1:])

    return output_list


#  insert many values into vfm db.
#  To be further abstracted for all dbs.
def insert_many_vfm_db(db_name, quarter, vfm_list):
    conn = sqlite3.connect(db_name + '.db')
    c = conn.cursor()
    c.executemany("INSERT INTO '{table}' VALUES (?,?,?,?,?,?,?,?)".format(table=quarter), vfm_list)
    conn.commit()
    conn.close()


#  returns a list of project names
def get_project_names(db_path, quarter):
    conn = sqlite3.connect(db_path)
    conn.row_factory = lambda cursor, row: row[0]
    c = conn.cursor()
    names = c.execute("SELECT project_name FROM '{table}'".format(table=quarter)).fetchall()
    conn.commit()
    conn.close()
    return names


#  Converts a db into a python dictionary when give a db and qrt list.
def convert_db_python_dict(db_path, quarter_list):
    conn = sqlite3.connect(db_path)

    # This is the important part, here we are setting row_factory property of
    # connection object to sqlite3.Row(sqlite3.Row is an implementation of
    # row_factory)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    output_dict = {}
    for quarter in quarter_list:
        inner_dict = {}
        project_names = get_project_names(db_path, quarter)
        for project in project_names:
            c.execute("select * from '{table}' WHERE project_name = '{p}'".format(table=quarter, p=project))
            result = [dict(row) for row in c.fetchall()][0]  # [0] there as output is dict in a list
            inner_dict[project] = result

        output_dict[quarter] = inner_dict

    conn.commit()
    conn.close()

    return output_dict
