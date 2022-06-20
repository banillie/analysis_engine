import configparser
import datetime
from pathlib import Path
from typing import Dict, List
from openpyxl import load_workbook, Workbook

from datamaps.api import project_data_from_master

from analysis_engine.data import root_path


def get_integration_data(
    confi_path: Path,
) -> Dict:
    # Returns a list of dft groups
    try:
        config = configparser.ConfigParser()
        config.read(confi_path)
        path_dict = {
            'project_map_path': config["GMPP INTEGRATION"]["project_map"],
            'gmpp_data_path': config["GMPP INTEGRATION"]["gmpp_data"],
            'key_map_path': config["GMPP INTEGRATION"]["key_map"],
            'master_comp_path': config["GMPP INTEGRATION"]["master_for_comparison"],
        }
    except:
        logger.critical("Configuration file issue. Please check and make sure it's correct.")
        sys.exit(1)

    return path_dict


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
                ipa_key = ipa_key.replace(',', '')
            if not gaps:
                ipa_key = ipa_key.replace('  ', ' ')
            if flip:
                output_dict[ipa_key] = dft_key
            else:
                output_dict[dft_key] = ipa_key

    return output_dict


def get_gmpp_data(
        file_name: str,
        # file_name_two: str
):
    from datetime import datetime
    import xlrd

    try:
        wb = load_workbook(root_path / 'input/{}.xlsx'.format(file_name), data_only=True)
    except: # bit of flexibility to help user with different file types
        wb = load_workbook(root_path / 'input/{}.xlsm'.format(file_name), data_only=True)
    ws = wb.active
    ws_list = [ws]

    key_map = get_map(load_workbook(root_path / "input/GMPP_INTEGRATION_KEY_MAP.xlsx"), commas=True, gaps=True)

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
        km_file_name: str,
        ipdc_d_file_path,
        project_list=False,
    ):
    wb = Workbook()
    ws = wb.active

    # get non-gmpp keys from team

    key_map = get_map(load_workbook(
        root_path / "input/{}.xlsx".format(km_file_name)
    ), commas=True, gaps=True)
    ipdc_data = project_data_from_master(root_path / "core_data/{}.xlsx".format(ipdc_d_file_path), 2, 2021)

    a_proj_name = ipdc_data.projects[1]
    # a_proj_name = 'East Coast Mainline Programme'
    if project_list:
        list_of_projects = project_list
    else:
        list_of_projects = list(initial_dict.keys())

    for x, project in enumerate(list_of_projects):
        ws.cell(row=1, column=3 + x).value = project
        i = 0
        for v in ipdc_data.data[a_proj_name].keys():
            # for ipa_key, dft_key in key_map.items():
            ws.cell(row=2 + i, column=1).value = v
            try:
                ipa_key = key_map[v]
                ws.cell(row=2 + i, column=2).value = ipa_key
                # try:
                ipa_value = initial_dict[project][ipa_key]
                if isinstance(ipa_value, datetime.datetime):
                    ipa_value = ipa_value.date()
                    ws.cell(row=2 + i, column=3 + x, value=ipa_value).number_format = "dd/mm/yy"
                ws.cell(row=2 + i, column=3 + x).value = ipa_value
                # except KeyError:
                #     pass
            except KeyError:
                pass
            i += 1

    ws.cell(row=1, column=1).value = "Project Name (DfT Keys)"
    ws.cell(row=1, column=2).value = "Project Name (IPA Keys)"

    wb.save(root_path / "output/gmpp_online_data_dft_master_format.xlsx")

