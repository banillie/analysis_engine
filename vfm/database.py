import sqlite3
from datamaps.api import project_data_from_master
import re


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


#  put master data into dB via python dictionary.
def import_master_to_db(db_path, master_path):
    m = project_data_from_master(master_path, 4, 2019)
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("INSERT INTO quarter VALUES ('{q}', '{q_int}')".
              format(q=m.quarter, q_int=m.quarter.quarter))
    c.execute("INSERT INTO milestone_type VALUES ('Approval', 'BLAH BLAH')")
    for project in m.projects:
        c.execute("INSERT INTO project_group VALUES ('{pg}')".
                  format(pg=m.data[project]['DfT Group']))
        c.execute("INSERT INTO project VALUES ('{q}', "
                  "'{pg}', '{pi}', '{p}')".format(
                    q=m.quarter,
                    pg=m.data[project]['DfT Group'],
                    pi=m.data[project]['DFT ID Number'],
                    p=project))
        for i in range(1, 2):
            m_type = "Approval MM" + str(i)
            if m_type in list(m.data[project].keys()):
                #  note string amended to remove ' and replace with `
                n = m.data[project]["Approval MM" + str(i) + " Notes"]
                note = n.replace("'", "`")
                c.execute(
                    f"""INSERT INTO milestone VALUES (
                    'Approval', 
                    '{m.data[project]['DFT ID Number']}',
                    '{project}', '{m.data[project]["Approval MM" + str(i)]}', 
                    '{m.data[project]["Approval MM" + str(i) + " Gov Type"]}',
                    '{m.data[project]["Approval MM" + str(i) + " Ver No"]}', 
                    '{m.data[project]["Approval MM" + str(i) + " Original Baseline"]}', 
                    '{m.data[project]["Approval MM" + str(i) + " Forecast / Actual"]}', 
                    '{m.data[project]["Approval MM" + str(i) + " Variance"]}', 
                    '{m.data[project]["Approval MM" + str(i) + " Status"]}', 
                    '{note}')""")

    conn.commit()


#  create master dB.
def create_db(db_path):
    conn = sqlite3.connect(db_path)
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
            (name text,
            quarter_number integer)""")

    c.execute("""CREATE TABLE 'project_group'
                (name text)""")

    c.execute("""CREATE TABLE 'project'
            (quarter_id integer,
            group_id integer,
            project_id integer,
            name text,
            FOREIGN KEY(quarter_id) REFERENCES quarter(id)
            FOREIGN KEY(group_id) REFERENCES project_group(id))""")

    c.execute("""CREATE TABLE 'milestone_type'
            (type text,
            description text)""")

    c.execute("""CREATE TABLE 'milestone'
            (milestone_type_id text,
            project_id integer,
            project_name text,
            name text,
            gov_type text,
            ver_no real,
            orig_baseline text,
            forecast_actual text,
            variance real,
            status text,
            notes text,
            FOREIGN KEY(project_id) REFERENCES project(quarter_id)
            FOREIGN KEY(milestone_type_id) REFERENCES milestone_type(id)
            FOREIGN KEY(project_name) REFERENCES project(name)
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


#  for querying db in python
def query_db(db_path, key, quarter):
    conn = sqlite3.connect(db_path)
    conn.row_factory = lambda cursor, row: row[0]
    c = conn.cursor()
    c.execute("SELECT {key} FROM {table}".format(key=key, table=quarter))
    result = c.fetchall()

    conn.commit()
    conn.close()

    return result


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
