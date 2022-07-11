import datetime
from typing import List, Union, Dict
import numpy as np

from analysis_engine.error_msgs import (
    logger,
    ProjectNameError,
    not_recognised_project_group_or_stage,
    not_recognised_quarter,
)

from analysis_engine.dictionaries import BC_STAGE_DICT


def get_iter_list(md, **kwargs) -> List[str]:
    # If quarter arg not passed in latest qrt return as default.
    if 'quarter' in kwargs:
        if kwargs['quarter'] == "standard":
            iter_list = [
                md['quarter_list'][0],
                md['quarter_list'][1],
            ]
        elif kwargs['quarter'] == ["all"]:
            iter_list = md['quarter_list']
        else:
            iter_list = kwargs['quarter']
    else:
        iter_list = [md['current_quarter']]

    return iter_list


def cal_group(
    group: List[str] or List[List[str]],
    report_type,
    md,
    tp_indx: int,
) -> List[str]:
    error_case = []
    output = []
    q_str = md["quarter_list"][tp_indx]  # quarter string
    for g in group:  # pg is project/group
        if g == "pipeline":
            continue
        try:
            local_g = md["dft_group"][q_str][g]
            output += local_g
        except KeyError:
            try:
                output.append(md["abbreviations"][g]["full name"])
            except KeyError:
                # try:
                #     output.append(md["full_names"][g])
                # except KeyError:
                try:
                    local_g = md["dft_group"][q_str][BC_STAGE_DICT[report_type][g]]
                    output += local_g
                except KeyError:
                    try:
                        local_g = md["dft_group"][q_str][g]
                        output += local_g
                    except KeyError:
                        error_case.append(g)

    not_recognised_project_group_or_stage(error_case)

    return output


def get_group(md, tp: str, class_kwargs) -> List[str]:
    try:
        tp_indx = md["quarter_list"].index(tp)
    except ValueError:
        not_recognised_quarter(tp)

    if "stage" in class_kwargs:
        group_list = class_kwargs["stage"]
    elif "group" in class_kwargs:
        group_list = class_kwargs["group"]
    else:
        group_list = md["groups"]

    group = cal_group(group_list, class_kwargs["report"], md, tp_indx)

    if "remove" in class_kwargs:
        group = remove_from_group(
            group, class_kwargs["remove"], md, tp_indx, class_kwargs
        )
    return group


def remove_from_group(
    pg_list: List[str],
    remove_list: List[str] or List[list[str]],
    master,
    tp_index: int,
) -> List[str]:
    if any(isinstance(x, list) for x in remove_list):
        remove_list = [item for sublist in remove_list for item in sublist]
    else:
        remove_list = remove_list
    removed_case = []
    q_str = master.quarter_list[tp_index]
    for pg in remove_list:
        try:
            local_g = master.project_stage[q_str][pg]
            pg_list = [x for x in pg_list if x not in local_g]
            removed_case.append(pg)
        except KeyError:
            try:
                local_g = master.meta_groupings[q_str][pg]
                pg_list = [x for x in pg_list if x not in local_g]
                removed_case.append(pg)
            except KeyError:
                try:
                    pg_list.remove(master.abbreviations[pg]["full name"])
                    removed_case.append(pg)
                except (ValueError, KeyError):
                    try:
                        pg_list.remove(master.full_names[pg])
                        removed_case.append(pg)
                    except (ValueError, KeyError):
                        pass

    if removed_case:
        for p in removed_case:
            logger.info(p + " successfully removed from analysis.")

    return pg_list


def get_correct_p_data(
    master,
    project_name: str,
    quarter: str,
) -> Dict[str, Union[str, int, datetime.date, float]]:
    for md in master['master_data']:
        if md['quarter'] == quarter:
            return md["data"][project_name]


def get_quarter_index(md, tp):
    return md["quarter_list"].index(tp)


def calculate_profiled(
    p: int or List[int], s: int or List[int], unpro: int or List[int]
) -> list:
    """small helper function to calculate the proper profiled amount. This is necessary as
    other wise 'profiled' would actually be the total figure.
    p = profiled list
    s = spent list
    unpro = unprofiled list"""
    if isinstance(p, list):
        f_profiled = []
        for y, amount in enumerate(p):
            t = amount - (s[y] + unpro[y])
            f_profiled.append(t)
        return f_profiled
    else:
        return p - (s + unpro)


def moving_average(x, w):
    return np.convolve(x, np.ones(w), "valid") / w
