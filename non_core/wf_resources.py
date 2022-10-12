from openpyxl import load_workbook, Workbook
from datamaps.api import project_data_from_master


def create_keys():
    """
    Quick way to iterate the numbers in front of keys with the same names
    """
    wb = load_workbook("/home/will/Documents/ipdc/input/dm_project_wf.xlsx")
    ws = wb.active
    int_list = []
    for cell in ws["A"]:
        if "1" in cell.value:
            int_list.append(cell.value)

    final_list = []
    for x in range(1, 35):
        for i in int_list:
            w = list(i)
            w[0] = str(x)
            final_list.append("".join(w))

    for i, x in enumerate(final_list):
        ws.cell(row=7 + i, column=1).value = x

    wb.save("/home/will/Documents/ipdc/input/dm_project_wf_altered.xlsx")


ks = [
    "1.1 Post (Job Title)",
    "1.2 Function",
    "1.3 Post FTE on this project ",
    "1.4 FTE less than one",
    "1.5 Project delivery professional",
    "1.6 How is the post resourced",
    "1.7 Is this post currently vacant",
    "1.8 Current post holder joining date",
    "1.9 Current post holder expected depart date",
    "1.10 Comments",
]


def data_for_pbi(master_path, save_path):
    wf_dict = project_data_from_master(master_path, 1, 2000)
    wb = Workbook()
    ws = wb.active
    chopped_dict = {}
    for project in wf_dict.projects:
        values = []
        for no in range(1, 26):
            for i in ks:
                word = list(i)
                word[0] = str(no)
                altered_key = "".join(word).rstrip()  # some keys have a trailing blank space.
                val = wf_dict[project][altered_key]
                if val is not None:
                    values.append((altered_key, val))
        chopped_dict[project] = dict(values)

    for project in wf_dict.projects:
        for no in range(1, 26):
            if no == 1:
                ws.cell(row=1, column=1).value = "Project"
            ws.cell(row=no + 1, column=1).value = project
            for i, key in enumerate(ks):
                if no == 1:
                    word = list(key)
                    ws.cell(row=1, column=2 + i).value = "".join(word[4:]).lstrip()
                word = list(key)
                word[0] = str(no)
                altered_key = "".join(word).rstrip()
                try:
                    val = chopped_dict[project][altered_key]
                    ws.cell(row=no + 1, column=2 + i).value = val
                except KeyError:
                    pass

    wb.save(save_path)

m_path = "/home/will/Documents/datamaps/output/wf_data_master_pilot.xlsx"
s_path = "/home/will/Documents/ipdc/output/wf_master_altered.xlsx"

data_for_pbi(m_path, s_path)