import datetime
from typing import List, Union, Dict
import numpy as np

from analysis_engine.error_msgs import (
    logger,
    ProjectNameError,
    not_recognised_project_group_or_stage,
    not_recognised_quarter,
)

from analysis_engine.dictionaries import (
    BC_STAGE_DICT_ABB_TO_FULL,
    BC_STAGE_DICT_FULL_TO_ABB,
)


def get_iter_list(md, **kwargs) -> List[str]:
    # If quarter arg not passed in latest qrt return as default.
    if "quarter" in kwargs:
        if kwargs["quarter"] == "standard":
            iter_list = [
                md["quarter_list"][0],
                md["quarter_list"][1],
            ]
        elif kwargs["quarter"] == "four":
            try:
                iter_list = [
                    md["quarter_list"][0],
                    md["quarter_list"][1],
                    md["quarter_list"][2],
                    md["quarter_list"][3],
                ]
            except:  # what is the exception
                iter_list = md["quarter_list"]
                logger.info("Less than four quarters present in core_data folder")

        elif kwargs["quarter"] == "all":
            iter_list = md["quarter_list"]
        else:
            iter_list = kwargs["quarter"]
    # this is the main default. one quarter. used for dandelion, milestones.
    else:
        iter_list = [md["current_quarter"]]

    return iter_list


def get_group(md, tp, **kwargs) -> List[str]:
    """
    A integral function for much of sorting of data. It groups projects in lists. These lists
    are then used to compile analysis. A project is grouped according to either its departmental
    group, or its business case stage.

    A complicating factor is that the group or stage a project is in/at could change between quarters
    and it is useful to preserve the status at each quarter. This is done via the meta_groupings
    which is calculated at the data initiate stage in PythonData.
    """
    initial_error_case = []
    output_list = []
    meta_groupings = md["meta_groupings"]
    abbreviations = md["abbreviations"]
    # need to standardise so always full name used.
    if "group" in kwargs:
        for g in kwargs["group"]:
            if g == "pipeline":
                continue
            try:
                loop_list = meta_groupings[tp][g]
                output_list += loop_list
            except KeyError:
                initial_error_case.append(g)

    if "stage" in kwargs:
        # inelegant loop there are two try statements to handle abbrevations
        # of stage terms e.g. Full Business Case and FBC. But must be a better
        # way to handle.
        for s in kwargs["stage"]:
            if s == "pipeline":
                continue
            try:  # full term
                loop_list = meta_groupings[tp][s]
                output_list += loop_list
            except KeyError:
                try:  # abbreviated term
                    loop_list = meta_groupings[tp][BC_STAGE_DICT_ABB_TO_FULL[s]]
                    output_list += loop_list
                except KeyError:
                    initial_error_case.append(s)

    final_error_case = []
    if initial_error_case:
        project_names = md[
            "project_information"
        ].keys()  # is this needed if just single project?
        for p in initial_error_case:
            if p in project_names:
                if p in md["master_data"][md["quarter_list"].index(tp)]["data"].keys():
                    output_list.append(p)
            elif p in list(abbreviations.keys()):
                pfn = abbreviations[
                    p
                ]  # pfn = project full name. coverts abbreviations back to full names
                if (
                    pfn
                    in md["master_data"][md["quarter_list"].index(tp)]["data"].keys()
                ):
                    output_list.append(pfn)
            else:
                final_error_case.append(p)

    if len(final_error_case) != 0:
        qrt_list = []
        for m in md["master_data"]:
            qrt_list.append(m["quarter"])

        if tp not in qrt_list:
            not_recognised_quarter(tp)
        else:
            not_recognised_project_group_or_stage(final_error_case)

    if "remove" in kwargs:
        for p in kwargs["remove"]:
            try:
                if p in output_list:
                    output_list.remove(p)
                if abbreviations[p] in output_list:
                    output_list.remove(abbreviations[p])
            except KeyError:
                final_error_case.append(p)
                not_recognised_project_group_or_stage(final_error_case)

    return output_list


# def remove_from_group(
#     pg_list: List[str],
#     remove_list: List[str] or List[list[str]],
#     master,
#     tp_index: int,
# ) -> List[str]:
#     if any(isinstance(x, list) for x in remove_list):
#         remove_list = [item for sublist in remove_list for item in sublist]
#     else:
#         remove_list = remove_list
#     removed_case = []
#     q_str = master.quarter_list[tp_index]
#     for pg in remove_list:
#         try:
#             local_g = master.project_stage[q_str][pg]
#             pg_list = [x for x in pg_list if x not in local_g]
#             removed_case.append(pg)
#         except KeyError:
#             try:
#                 local_g = master.meta_groupings[q_str][pg]
#                 pg_list = [x for x in pg_list if x not in local_g]
#                 removed_case.append(pg)
#             except KeyError:
#                 try:
#                     pg_list.remove(master.abbreviations[pg]["full name"])
#                     removed_case.append(pg)
#                 except (ValueError, KeyError):
#                     try:
#                         pg_list.remove(master.full_names[pg])
#                         removed_case.append(pg)
#                     except (ValueError, KeyError):
#                         pass
#
#     if removed_case:
#         for p in removed_case:
#             logger.info(p + " successfully removed from analysis.")
#
#     return pg_list
#


def get_correct_p_data(
    master,
    project_name: str,
    quarter: str,
) -> Dict[str, Union[str, int, datetime.date, float]]:
    for md in master["master_data"]:
        if md["quarter"] == quarter:
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
