from typing import List

from analysis_engine.error_msgs import logger, ProjectNameError, not_recognised_project_or_group


def get_iter_list(class_kwargs, master) -> List[str]:
    iter_list = []
    if "baseline" in class_kwargs:
        if class_kwargs["baseline"] == ["standard"]:
            iter_list = ["current", "last", "bl_one"]
        elif class_kwargs["baseline"] == ["all"]:
            iter_list = ["current", "last", "bl_one", "bl_two", "bl_three"]
        elif class_kwargs["baseline"] == ["standard"]:
            iter_list = ["current", "last", "bl_one"]
        else:
            iter_list = class_kwargs["baseline"]

    elif "quarter" in class_kwargs:
        if class_kwargs["quarter"] == ["standard"]:
            iter_list = [
                master.quarter_list[0],
                master.quarter_list[1],
            ]
        elif class_kwargs["quarter"] == ["all"]:
            iter_list = master.quarter_list
        else:
            iter_list = class_kwargs["quarter"]

    return iter_list


def cal_group(input_list: List[str] or List[List[str]], md, tp_indx: int,
        # input_list_indx=None,
) -> List[str]:
    """
    What does this do?
    """

    error_case = []
    output = []
    # if input_list_indx or input_list_indx == 0:
    #     input_list = [input_list[input_list_indx]]
    # if any(isinstance(x, list) for x in input_list):
    #     inner_list = [item for sublist in input_list for item in sublist]
    # else:
    #     inner_list = input_list
    q_str = md['quarter_list'][tp_indx]  # quarter string
    for g in input_list:  # pg is project/group
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
    ##  baselines not in use
    # if "baseline" in class_kwargs:
    #     tp_indx = 0  # baseline uses latest project group only
    # elif "quarter" in class_kwargs:

    tp_indx = md['quarter_list'].index(tp)

    if "stage" in class_kwargs:
        # if group_indx or group_indx == 0:
        #     group = cal_group(class_kwargs["stage"], md, tp_indx, group_indx)
        # else:
        group = cal_group(class_kwargs["stage"], md, tp_indx)
    elif "group" in class_kwargs:
        # if group_indx or group_indx == 0:
        #     group = cal_group(class_kwargs["group"], md, tp_indx, group_indx)
        # else:
        group = cal_group(class_kwargs["group"], md, tp_indx)
    else:
        group = cal_group(md.current_projects, md, tp_indx)

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
        c_kwargs,  # class_kwargs
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

        ## not sure why quarter / baseline is important hashing for now.
        # if "baseline" in c_kwargs:
        #     for p in removed_case:
        #         logger.critical(p + " not a recognised.")
        #     raise ProjectNameError(
        #         'Program stopping. Please check the "remove" entry and re-enter.'
        #     )
        # if "quarter" in c_kwargs:
        #     for p in removed_case:
        #         logger.info(
        #             p + " not a recognised or not present in " + q_str + "."
        #             '"So not removed from the data for that quarter. Make sure the '
        #             '"remove" entry is correct.'
        #         )

    return pg_list