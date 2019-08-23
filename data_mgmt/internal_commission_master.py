from openpyxl import load_workbook

def get_key_names(workbook):
    key_list = []
    lst_q_key_list = []
    ws = workbook.active

    for row_num in range(2, ws.max_row + 1):
        key_list.append(ws.cell(row=row_num, column=1).value)

    for key in key_list:
        lst_q_key_list.append('lst qrt ' + str(key))

    return lst_q_key_list


wb = load_workbook(
        'C:\\Users\\Standalone\\general\\masters folder\\core_data\\master_1_2019_wip_(25_7_19).xlsx')

test = get_key_names(wb)