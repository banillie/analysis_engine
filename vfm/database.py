import sqlite3


def create_connect_db(db_name):
    conn = sqlite3.connect(db_name + '.db')
    return conn

#  create a new table in vfm db.
def create_vfm_table(conn, insert_quarter):
    # conn = sqlite3.connect(db_name + '.db')
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


def insert_many_vfm_db(conn, quarter, vfm_list):
    c = conn.cursor()
    c.executemany("INSERT INTO '{table}' VALUES (?,?,?,?,?,?,?,?)".format(table=quarter), vfm_list)
    #c.executemany('INSERT INTO q4_1920 VALUES (?,?,?,?,?,?,?,?)', vfm_q4_1920)

    conn.commit()
    conn.close()