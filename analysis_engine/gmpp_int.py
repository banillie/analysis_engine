import datetime
from typing import Dict
from openpyxl import load_workbook, Workbook

from datamaps.api import project_data_from_master
from analysis_engine.settings import get_integration_data


def get_map(wb, commas=False, gaps=False, flip=False):
    ws = wb.active
    output_dict = {}

    for x in range(2, ws.max_row + 1):
        ipa_key = ws.cell(row=x, column=2).value
        if ipa_key in output_dict.keys():
            pass
        if ipa_key is None:
            pass
        else:
            dft_key = ws.cell(row=x, column=1).value
            ipa_key = ws.cell(row=x, column=2).value
            if not commas:
                ipa_key = ipa_key.replace(",", "")
            if not gaps:
                ipa_key = ipa_key.replace("  ", " ")
            if flip:
                output_dict[ipa_key] = dft_key
            else:
                output_dict[dft_key] = ipa_key

    return output_dict


def get_gmpp_online_data(**op_args):
    from datetime import datetime
    import xlrd

    file_name = op_args['gmpp_data_path']
    try:
        wb = load_workbook(op_args['root_path'] + '/input/{}.xlsx'.format(file_name, data_only=True))
    except:
        wb = load_workbook(op_args['root_path'] + '/input/{}.xlsm'.format(file_name, data_only=True))

    ws = wb.active
    ws_list = [ws]

    key_map = get_map(
        load_workbook(op_args['root_path'] + "/input/GMPP_INTEGRATION_KEY_MAP.xlsx"),
        commas=True,
        gaps=True,
    )

    initial_dict = {}
    missing_keys = []
    for ws in ws_list:
        for x in range(24, ws.max_row + 1):
            project_name = ws.cell(row=x, column=2).value
            key = ws.cell(row=x, column=6).value
            if key not in key_map.values():
                if key not in missing_keys:
                    missing_keys.append(key)
            s_value = ws.cell(row=x, column=7).value
            n_value = ws.cell(row=x, column=8).value
            if n_value != 0:
                s_value = n_value
            if "Date" in key or "date" in key or "6.03c: To" in key:
                # s_value = datetime(*xlrd.xldate_as_tuple(n_value, 0))
                # print(project_name, key, n_value)
                if n_value > 20000:
                    # if "7.02.10" in key:
                    #     pass
                    # else:
                    s_value = datetime(*xlrd.xldate_as_tuple(n_value, 0))
                    # else:
                    #     s_value = n_value
            if "Grade" in key:  # to make grade 6 consistent with dft data
                try:
                    s_value = int(s_value)
                except ValueError:
                    pass

            if project_name in list(initial_dict.keys()):
                initial_dict[project_name][key] = s_value
            else:
                initial_dict[project_name] = {key: s_value}

    ## Use for checking to see if key map missing keys.
    # for x in missing_keys:
    #     print(x)

    return initial_dict


def place_gmpp_online_keys_into_dft_master_format(
    initial_dict: Dict,
    # km_file_name: str,
    # ipdc_d_file_path,
    project_list=False,
    **op_args,
):
    wb = Workbook()
    ws = wb.active

    # get non-gmpp keys from team

    key_map = get_map(
        load_workbook(op_args['root_path'] + "/input/{}.xlsx".format(op_args["key_map_path"])),
        commas=True,
        gaps=True,
    )
    # ipdc_data = project_data_from_master(
    #     op_args['root_path'] + "/core_data/{}.xlsx".format(op_args["master_comp_path"]), 2, 2021
    # )
    project_map = get_map(
        load_workbook(op_args['root_path'] + '/input/GMPP_INTEGRATION_PROJECT_MAP.xlsx')
    )


    # a_proj_name = ipdc_data.projects[1]
    # # a_proj_name = 'East Coast Mainline Programme'
    # if project_list:
    #     list_of_projects = project_list
    # else:
    list_of_projects = list(initial_dict.keys())

    for x, project in enumerate(list_of_projects):
        ws.cell(row=1, column=3 + x).value = project_map[project]
        i = 0
        for i, k in enumerate(initial_dict[project].keys()):
            v = initial_dict[project][k]
            ws.cell(row=2 + i, column=3 + x).value = v
            if isinstance(v, datetime.datetime):
                ws.cell(row=2 + i, column=3 + x).number_format = "dd/mm/yy"
        # for v in ipdc_data.data[a_proj_name].keys():
        #     # for ipa_key, dft_key in key_map.items():
        #     ws.cell(row=2 + i, column=1).value = v
        #     try:
        #         ipa_key = key_map[v]
        #         ws.cell(row=2 + i, column=2).value = ipa_key
        #         # try:
        #         ipa_value = initial_dict[project][ipa_key]
        #         if isinstance(ipa_value, datetime.datetime):
        #             ipa_value = ipa_value.date()
        #             ws.cell(
        #                 row=2 + i, column=3 + x, value=ipa_value
        #             ).number_format = "dd/mm/yy"
        #         ws.cell(row=2 + i, column=3 + x).value = ipa_value
        #         # except KeyError:
        #         #     pass
        #     except KeyError:
        #         pass
        #     i += 1

    ws.cell(row=1, column=1).value = "Project Name (DfT Keys)"
    ws.cell(row=1, column=2).value = "Project Name (IPA Keys)"

    wb.save(op_args['root_path'] + "/output/gmpp_online_data_dft_master_format.xlsx")


def get_gmpp_data(**op_args):
    get_integration_data(op_args)
    gmpp_data = get_gmpp_online_data(**op_args)
    place_gmpp_online_keys_into_dft_master_format(gmpp_data, **op_args,)


## Code not currently in use.
# def data_check_print_out(
#         ipdc_d_file_path: str,
#         km_file_name: str,
#         pn_file_name: str,
# ):
#     gmpp_data = project_data_from_master(root_path / "input/gmpp_online_data_temp.xlsx", 2, 2021)
#     os.remove(root_path / "input/gmpp_online_data_temp.xlsx")
#     ipdc_data = project_data_from_master(root_path / "core_data/{}.xlsx".format(ipdc_d_file_path), 2, 2021)
#     key_map = get_map(load_workbook
#                       (root_path / "input/{}.xlsx".format(km_file_name)), flip=True)
#     project_map = get_map(load_workbook
#                           (root_path / "input/{}.xlsx".format(pn_file_name)))
#
#     wb = Workbook()
#     ws = wb.active
#
#     def remove_keys(key):
#         output = key
#         for rk in RK_LIST:  # remove key
#             if rk in key:
#                 output = "remove"
#         return output
#
#     start_row = 2
#     project_check_list = []
#     for x, project in enumerate(list(project_map.keys())):  # could be project_map.keys()
#         project_check_list.append(project)
#         try:  # exception so only projects in ipdc data compared.
#             for i, k in enumerate(gmpp_data.data[project]):
#                 if k is None:
#                     continue
#                 check_key = remove_keys(k)
#                 if check_key == "remove":
#                     continue
#                 ws.cell(row=start_row, column=1).value = project
#                 try:
#                     dft_project_name = project_map[project]
#                     project_check = "PASS"
#                 except KeyError:
#                     dft_project_name = ""
#                     project_check = "FAILED"
#                 ws.cell(row=start_row, column=2).value = dft_project_name
#                 ws.cell(row=start_row, column=3).value = project_check
#                 ws.cell(row=start_row, column=4).value = k
#                 try:
#                     dft_key_name = key_map[k]
#                     if dft_key_name == "None":
#                         continue
#                     key_check = "PASS"
#                 except KeyError:
#                     # print(k)
#                     dft_key_name = ""
#                     key_check = "FAILED"
#                 ws.cell(row=start_row, column=5).value = dft_key_name
#                 ws.cell(row=start_row, column=6).value = key_check
#
#                 gmpp_val = gmpp_data[project][k]
#
#                 try:
#                     dft_val = ipdc_data[dft_project_name][dft_key_name]
#                     if "Ver No" in dft_key_name or "Version No" in dft_key_name:
#                         if dft_val is not None:
#                             dft_val = str(dft_val)
#                         if gmpp_val is not None:
#                             gmpp_val = str(gmpp_val)
#                     # if 'Phone No' in dft_key_name:  # started to think about tele nos but leaving for now.
#                     #     print(gmpp_val)
#                     #     print(dft_val)
#                 except KeyError:
#                     dft_val = ""
#
#                 ws.cell(row=start_row, column=7).value = gmpp_val
#                 if isinstance(gmpp_val, datetime.datetime):
#                     gmpp_val = gmpp_val.date()
#                     ws.cell(row=start_row, column=7, value=gmpp_val).number_format = "dd/mm/yy"
#
#                 ws.cell(row=start_row, column=8).value = dft_val
#                 if isinstance(dft_val, datetime.datetime):
#                     dft_val = dft_val.date()
#                     ws.cell(row=start_row, column=8, value=dft_val).number_format = "dd/mm/yy"
#
#                 if gmpp_val in list(GMPP_M_DICT.keys()):
#                     if GMPP_M_DICT[gmpp_val] == dft_val:
#                         ws.cell(row=start_row, column=9).value = "MATCH"
#                         start_row += 1
#                         continue
#
#                 if isinstance(gmpp_val, str) and isinstance(dft_val, str):
#                     if "Ver No" in dft_key_name or "Version No" in k:
#                         try:
#                             gmpp_val = int(float(gmpp_val))
#                             dft_val = int(float(dft_val))
#                         except ValueError:
#                             pass
#                     else:
#                         gmpp_val = gmpp_val.split()
#                         dft_val = dft_val.split()
#
#                 # get floats of different lengths to match
#                 if isinstance(dft_val, float) and isinstance(gmpp_val, float):
#                     dft_val = float("{:.2f}".format(dft_val))
#                     gmpp_val = float("{:.2f}".format(gmpp_val))
#
#                 if isinstance(dft_val, float) and isinstance(gmpp_val, int):
#                     dft_val = round(dft_val)
#
#                 if isinstance(dft_val, int) and isinstance(gmpp_val, float):
#                     gmpp_val = round(gmpp_val)
#
#                 if gmpp_val == dft_val:
#                     ws.cell(row=start_row, column=9).value = "MATCH"
#                 elif gmpp_val is None and dft_val == "":
#                     ws.cell(row=start_row, column=9).value = "MATCH"
#                 elif gmpp_val == "" and dft_val is None:
#                     ws.cell(row=start_row, column=9).value = "MATCH"
#                 elif gmpp_val is None and dft_val == 0:
#                     ws.cell(row=start_row, column=9).value = "MATCH"
#                 elif dft_key_name in IGNORE_LIST:
#                     # print(dft_key_name)
#                     ws.cell(row=start_row, column=9).value = "IGNORE"
#                 else:
#                     ws.cell(row=start_row, column=9).value = "DIFFERENT"
#
#                 start_row += 1
#         except KeyError:
#             pass
#     ws.cell(row=1, column=1).value = "GMPP PROJECT NAME"
#     ws.cell(row=1, column=2).value = "DFT PROJECT NAME"
#     ws.cell(row=1, column=3).value = "NAME CHECK"
#     ws.cell(row=1, column=4).value = "GMPP KEY"
#     ws.cell(row=1, column=5).value = "DFT KEY"
#     ws.cell(row=1, column=6).value = "KEY CHECK"
#     ws.cell(row=1, column=7).value = "GMPP VALUE"
#     ws.cell(row=1, column=8).value = "DFT VALUE"
#     ws.cell(row=1, column=9).value = "VALUE CHECK"
#
#     p_check = [x for x in list(project_map.keys()) if x not in project_check_list]
#     if not p_check:
#         pass
#     else:
#         print("note following projects missing:")
#         for x in p_check:
#             print(x)
#
#     wb.save(root_path / f"output/GMPP_IPDC_DATA_CHECK_USING_{ipdc_d_file_path}.xlsx")

