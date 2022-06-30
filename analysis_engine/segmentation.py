import datetime
from typing import List, Union, Dict

from analysis_engine.error_msgs import logger, ProjectNameError, not_recognised_project_or_group


def get_iter_list(class_kwargs, master) -> List[str]:
    iter_list = []
    if class_kwargs["quarter"] == ["standard"]:
        iter_list = [
            master['quarter_list'][0],
            master['quarter_list'][1],
        ]
    elif class_kwargs["quarter"] == ["all"]:
        iter_list = master['quarter_list']
    else:
        iter_list = class_kwargs["quarter"]

    return iter_list


def cal_group(group: List[str] or List[List[str]], md, tp_indx: int,
              # input_list_indx=None,
              ) -> List[str]:
    error_case = []
    output = []
    # if input_list_indx or input_list_indx == 0:
    #     input_list = [input_list[input_list_indx]]
    # if any(isinstance(x, list) for x in input_list):
    #     inner_list = [item for sublist in input_list for item in sublist]
    # else:
    #     inner_list = input_list
    q_str = md['quarter_list'][tp_indx]  # quarter string
    for g in group:  # pg is project/group
        if g == "pipeline":
            continue
        try:
            local_g = md['dft_group'][q_str][g]
            output += local_g
        # except KeyError:
        #     try:
        #         local_g = md['dft_group'][q_str][g]
        #         output += local_g
        except KeyError:
            try:
                output.append(md['abbreviations'][g]["full name"])
            except KeyError:
                try:
                    output.append(md['full_names'][g])
                except KeyError:
                    error_case.append(g)

    not_recognised_project_or_group(error_case)

    return output


def get_group(md, tp: str, class_kwargs, group_indx=None) -> List[str]:
    tp_indx = md['quarter_list'].index(tp)

    if "stage" in class_kwargs:
        group = cal_group(class_kwargs["stage"], md, tp_indx)
    elif "group" in class_kwargs:
        group = cal_group(class_kwargs["group"], md, tp_indx)
    else:
        # group = cal_group(md['current_projects'], md, tp_indx)  # why is this current_projects
        group = cal_group(md['groups'], md, tp_indx)  # why is this current_projects

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
        time_period: str,
) -> Dict[str, Union[str, int, datetime.date, float]]:
    tp_idx = master['quarter_list'].index(time_period)
    try:
        return master['master_data'][tp_idx]['data'][project_name]
    except KeyError: # KeyError handles project not reporting in quarter.
        return None


def get_quarter_index(md, tp):
    return md['quarter_list'].index(tp)
