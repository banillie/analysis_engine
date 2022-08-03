from analysis_engine.segmentation import (
    get_iter_list,
    get_group,
    get_correct_p_data,
    moving_average,
)
from analysis_engine.cleaning import convert_none_types
from analysis_engine.dictionaries import STANDARDISE_COST_KEYS
from analysis_engine.error_msgs import ProjectNameError, logger
from analysis_engine.settings import get_remove_income


class CostData:
    def __init__(self, master, **kwargs):
        self.master = master
        self.kwargs = kwargs
        self.report = kwargs["report"]
        self.quarters = self.master["quarter_list"]
        self.totals = {}
        self.baseline = {}
        self.get_totals()

    def get_totals(self) -> None:
        lower_dict = {}
        rm = get_remove_income(self.kwargs)
        for tp in self.quarters:
            spent = 0
            remaining = 0
            total = 0
            income_achieved = 0
            income_remaining = 0
            income_total = 0
            group = get_group(self.master, tp, **self.kwargs)
            for project_name in group:
                p_data = get_correct_p_data(self.master, project_name, tp)
                #  All try statements temporary until consistency between rpting datasets
                total += convert_none_types(
                    p_data[STANDARDISE_COST_KEYS[self.report]["total"]]
                )
                try:
                    spent += convert_none_types(
                        p_data[STANDARDISE_COST_KEYS[self.report]["spent"]]
                    )
                except KeyError:
                    spent += 0
                try:
                    income_achieved += convert_none_types(
                        p_data[STANDARDISE_COST_KEYS[self.report]["income_achieved"]]
                    )
                except KeyError:
                    income_achieved += 0
                try:
                    income_remaining += convert_none_types(
                        p_data[STANDARDISE_COST_KEYS[self.report]["income_remaining"]]
                    )
                except KeyError:
                    income_remaining += 0

                income_total += convert_none_types(
                    p_data[STANDARDISE_COST_KEYS[self.report]["income_total"]]
                )

                # option here handled via config file
                if project_name in rm:
                    try:
                        total - convert_none_types(
                            p_data[STANDARDISE_COST_KEYS[self.report]["income_total"]]
                        )
                        if [project_name] == group:
                            logger.info(
                                f"income has been removed from the total of {project_name}"
                            )
                    except KeyError:  # some older masters don't have key.
                        raise ProjectNameError()

            try:
                remaining += total - spent
            except KeyError:
                remaining += 0

            lower_dict[tp] = {
                "costs_spent": spent,
                "costs_remaining": remaining,
                "total": total,
                "income_achieved": income_achieved,
                "income_remaining": income_remaining,
                "income_total": income_total,
            }

        # if self.kwargs['report'] == "ipdc":
        #     for tp in self.iter_list:
        #         self.group = get_group(self.master, tp, self.kwargs)
        #         rdel_total = []
        #         cdel_total = []
        #         ngov_total = []
        #         total = []
        #         for project_name in self.group:
        #             p_data = get_correct_p_data(self.master, project_name, tp)
        #             if p_data is None:
        #                 break
        #             else:
        #                 rt = convert_none_types(p_data["Total RDEL Forecast Total"])
        #                 rdel_total.append(rt)
        #                 ct = convert_none_types(p_data["Total CDEL Forecast Total WLC"])
        #                 cdel_total.append(ct)
        #                 ng = convert_none_types(p_data["Non-Gov Total Forecast"])
        #                 ngov_total.append(ng)
        #                 t = convert_none_types(p_data["Total Forecast"])
        #                 # hard coded due to current use need.
        #
        #                 if project_name in self.kwargs['remove income from totals']:
        #                     try:
        #                         t = t - p_data[
        #                             "Total Forecast - Income both Revenue and Capital"
        #                         ]
        #                     except KeyError:  # some older masters do have key.
        #                         pass
        #                 total.append(t)
        #
        #             # rdel_profiled.append(rt - (rs + ru))
        #             # cdel_profiled.append(ct - (cs + cu))
        #             # profiled.append(t - (s + u))
        #
        #         lower_dict[tp] = {
        #             # "cat_spent": [sum(rdel_spent), sum(cdel_spent)],
        #             # "cat_prof": [sum(rdel_profiled), sum(cdel_profiled)],
        #             # "cat_unprof": [sum(rdel_unprofiled), sum(cdel_unprofiled)],
        #             # "spent": sum(spent),
        #             # "prof": sum(profiled),
        #             # "unprof": sum(unprofiled),
        #             "total": sum(total),
        #             "rdel": sum(rdel_total),
        #             "cdel": sum(cdel_total) - sum(ngov_total),
        #             "ngov": sum(ngov_total),
        #         }

        self.totals = lower_dict

    # def get_cost_profile(self) -> None:
    #     """Returns several lists which contain the sum of different cost profiles for the group of project
    #     contained with the master"""
    #     self.iter_list = get_iter_list(self.kwargs, self.master)
    #     lower_dict = {}
    #     for tp in self.iter_list:
    #         yearly_profile = []
    #         rdel_yearly_profile = []
    #         cdel_yearly_profile = []
    #         ngov_yearly_profile = []
    #         self.group = get_group(self.master, tp, self.kwargs)
    #         for year in YEAR_LIST:
    #             cost_total = 0
    #             rdel_total = 0
    #             cdel_total = 0
    #             ngov_total = 0
    #             for cost_type in COST_KEY_LIST:
    #                 for p in self.group:
    #                     p_data = get_correct_p_data(
    #                         self.kwargs, self.master, self.baseline_type, p, tp
    #                     )
    #                     if p_data is None:
    #                         continue
    #                     try:
    #                         cost = p_data[year + cost_type]
    #                         if cost is None:
    #                             cost = 0
    #                         cost_total += cost
    #                     except KeyError:  # handles data across different financial years via proj_info
    #                         try:
    #                             cost = self.master.project_information[p][
    #                                 year + cost_type
    #                                 ]
    #                         except KeyError:
    #                             cost = 0
    #                         if cost is None:
    #                             cost = 0
    #                         cost_total += cost
    #
    #                     if cost_type == COST_KEY_LIST[0]:  # rdel
    #                         rdel_total += cost
    #                     if cost_type == COST_KEY_LIST[1]:  # cdel
    #                         cdel_total += cost
    #                     if cost_type == COST_KEY_LIST[2]:  # ngov
    #                         ngov_total += cost
    #
    #             yearly_profile.append(cost_total)
    #             rdel_yearly_profile.append(rdel_total)
    #             cdel_yearly_profile.append(cdel_total)
    #             ngov_yearly_profile.append(ngov_total)
    #         lower_dict[tp] = {
    #             "prof": yearly_profile,
    #             "prof_ra": moving_average(yearly_profile, 2),
    #             "rdel": rdel_yearly_profile,
    #             "cdel": cdel_yearly_profile,
    #             "ngov": ngov_yearly_profile,
    #         }
    #     self.c_profiles = lower_dict
    #
    # def get_forecast_cost_profile(self) -> None:
    #     COST_CAT = [" RDEL ", " CDEL "]
    #     self.iter_list = get_iter_list(self.kwargs, self.master)
    #     tp_dict = {}
    #     for tp in self.iter_list:
    #         self.group = get_group(self.master, tp, self.kwargs)
    #         project_dict = {}
    #         list_total_total = []
    #         list_rdel_total = []
    #         list_cdel_total = []
    #         list_ngov_total = []
    #         list_std = []
    #         for p in self.group:
    #             RDEL_FORECAST_COST_KEYS = {
    #                 "Forecast one off new costs": [],
    #                 "Forecast recurring new costs": [],
    #                 "Forecast recurring old costs": [],
    #                 "Forecast Non Gov costs": [],
    #                 "Forecast Total": [],
    #                 "Forecast Income": [],
    #             }
    #             CDEL_FORECAST_COST_KEYS = {
    #                 "Forecast one off new costs": [],
    #                 "Forecast recurring new costs": [],
    #                 "Forecast recurring old costs": [],
    #                 " Forecast Non-Gov": [],
    #                 "Forecast Total WLC": [],
    #                 " Forecast - Income both Revenue and Capital": [],
    #             }
    #             p_data = get_correct_p_data(
    #                 self.kwargs, self.master, self.baseline_type, p, tp
    #             )
    #             if p_data is None:
    #                 continue
    #             rdel_std = convert_none_types(p_data["20-21 RDEL STD Total"])
    #             cdel_std = convert_none_types(p_data["20-21 CDEL STD Total"])
    #             list_std.append(rdel_std + cdel_std)
    #             for y in YEAR_LIST:
    #                 for cat in COST_CAT:
    #                     if cat == ' RDEL ':
    #                         for k in RDEL_FORECAST_COST_KEYS.keys():
    #                             if y in ["16-17", "17-18", "18-19"]:
    #                                 try:
    #                                     rdel = convert_none_types(self.master.project_information[p][y + cat + k])
    #                                 except KeyError:
    #                                     rdel = 0
    #                                     print(y + cat + k + " not found.")
    #                             else:
    #                                 rdel = convert_none_types(p_data[y + cat + k])
    #                             RDEL_FORECAST_COST_KEYS[k].append(rdel)
    #                     if cat == ' CDEL ':
    #                         for k in CDEL_FORECAST_COST_KEYS.keys():
    #                             if y in ["16-17", "17-18", "18-19"]:
    #                                 try:
    #                                     cdel = convert_none_types(self.master.project_information[p][y + cat + k])
    #                                 except KeyError:
    #                                     try:
    #                                         cdel = convert_none_types(self.master.project_information[p][y + k])
    #                                     except KeyError:
    #                                         cdel = 0
    #                                         print(y + k + " not found.")
    #                             else:
    #                                 try:
    #                                     cdel = convert_none_types(p_data[y + cat + k])
    #                                 except KeyError:
    #                                     try:
    #                                         cdel = convert_none_types(p_data[y + k])
    #                                     except KeyError:
    #                                         # user messaging if necessary
    #                                         # print(tp + " " + y + k + ' could not be found. Check')
    #                                         cdel = 0
    #                             CDEL_FORECAST_COST_KEYS[k].append(cdel)
    #
    #                 total_adding = [RDEL_FORECAST_COST_KEYS["Forecast Total"],
    #                                 CDEL_FORECAST_COST_KEYS["Forecast Total WLC"]]
    #                 year_total = [sum(x) for x in zip(*total_adding)]
    #                 ngov_adding = [RDEL_FORECAST_COST_KEYS["Forecast Non Gov costs"],
    #                                CDEL_FORECAST_COST_KEYS[" Forecast Non-Gov"]]
    #                 ngov_total = [sum(x) for x in zip(*ngov_adding)]
    #             # adding individual project data to dict is not necessary
    #             # project_dict[p] = {
    #             #     "rdel": RDEL_FORECAST_COST_KEYS["Forecast Total"],
    #             #     "cdel": CDEL_FORECAST_COST_KEYS["Forecast Total WLC"],
    #             #     "ngov": ngov_total,
    #             #     "total": year_total,
    #             # }
    #             list_total_total.append(year_total)
    #             list_ngov_total.append(ngov_total)
    #             list_rdel_total.append(RDEL_FORECAST_COST_KEYS["Forecast Total"])
    #             list_cdel_total.append(CDEL_FORECAST_COST_KEYS["Forecast Total WLC"])
    #
    #         project_dict["total"] = [sum(x) for x in zip(*list_total_total)]
    #         project_dict["rdel_total"] = [sum(x) for x in zip(*list_rdel_total)]
    #         project_dict["cdel_total"] = [sum(x) for x in zip(*list_cdel_total)]
    #         project_dict["ngov_total"] = [sum(x) for x in zip(*list_ngov_total)]
    #         project_dict["std"] = list_std
    #
    #         self.profiles[tp] = project_dict
    #
    # def get_baseline_cost_profile(self) -> None:
    #     COST_CAT = [" RDEL ", " CDEL "]
    #     # self.iter_list = get_iter_list(self.kwargs, self.master)
    #     self.iter_list = [self.master.current_quarter]
    #     tp_dict = {}
    #     for tp in self.iter_list:
    #         self.group = get_group(self.master, tp, self.kwargs)
    #         project_dict = {}
    #         list_total_total = []
    #         list_rdel_total = []
    #         list_cdel_total = []
    #         list_ngov_total = []
    #         for p in self.group:
    #             RDEL_BL_COST_KEYS = {
    #                 "BL one off new costs": [],
    #                 "BL recurring new costs": [],
    #                 "BL recurring old costs": [],
    #                 "BL Non Gov costs": [],
    #                 "BL Total": [],
    #                 "BL Income": [],
    #             }
    #             CDEL_BL_COST_KEYS = {
    #                 "BL one off new costs": [],
    #                 "BL recurring new costs": [],
    #                 "BL recurring old costs": [],
    #                 " BL Non-Gov": [],
    #                 "BL WLC": [],
    #                 " BL Income both Revenue and Capital": [],
    #             }
    #             # at moment bl only coming from current quarters data
    #             p_data = self.master.master_data[0]["data"][p]
    #             # p_data = get_correct_p_data(
    #             #     self.kwargs, self.master, self.baseline_type, p, tp
    #             # )
    #             if p_data is None:
    #                 continue
    #             for y in YEAR_LIST:
    #                 for cat in COST_CAT:
    #                     if cat == ' RDEL ':
    #                         for k in RDEL_BL_COST_KEYS.keys():
    #                             if y in ["16-17", "17-18", "18-19"]:
    #                                 rdel = 0
    #                             else:
    #                                 rdel = convert_none_types(p_data[y + cat + k])
    #                             RDEL_BL_COST_KEYS[k].append(rdel)
    #                     if cat == ' CDEL ':
    #                         for k in CDEL_BL_COST_KEYS.keys():
    #                             if y in ["16-17", "17-18", "18-19"]:
    #                                 cdel = 0
    #                             else:
    #                                 try:
    #                                     cdel = convert_none_types(p_data[y + cat + k])
    #                                 except KeyError:
    #                                     try:
    #                                         cdel = convert_none_types(p_data[y + k])
    #                                     except KeyError:
    #                                         # user messaging if necessary
    #                                         # print(tp + " " + y + k + ' could not be found. Check')
    #                                         cdel = 0
    #                             CDEL_BL_COST_KEYS[k].append(cdel)
    #                 total_adding = [RDEL_BL_COST_KEYS["BL Total"], CDEL_BL_COST_KEYS["BL WLC"]]
    #                 year_total = [sum(x) for x in zip(*total_adding)]
    #                 ngov_adding = [RDEL_BL_COST_KEYS["BL Non Gov costs"], CDEL_BL_COST_KEYS[" BL Non-Gov"]]
    #                 ngov_total = [sum(x) for x in zip(*ngov_adding)]
    #             # project_dict[p] = {
    #             #     "rdel": RDEL_BL_COST_KEYS["BL Total"],
    #             #     "cdel": CDEL_BL_COST_KEYS["BL WLC"],
    #             #     "total": year_total,
    #             # }
    #             list_total_total.append(year_total)
    #             list_ngov_total.append(ngov_total)
    #             list_rdel_total.append(RDEL_BL_COST_KEYS["BL Total"])
    #             list_cdel_total.append(CDEL_BL_COST_KEYS["BL WLC"])
    #         project_dict["total"] = [sum(x) for x in zip(*list_total_total)]
    #         project_dict["rdel_total"] = [sum(x) for x in zip(*list_rdel_total)]
    #         project_dict["cdel_total"] = [sum(x) for x in zip(*list_cdel_total)]
    #         project_dict["ngov_total"] = [sum(x) for x in zip(*list_ngov_total)]
    #         # tp_dict["baseline"] = project_dict
    #         self.profiles["baseline"] = project_dict
    #
    # def get_wlc_data(self) -> None:
    #     """
    #     calculates the quarters total wlc change
    #     """
    #     self.iter_list = get_iter_list(self.kwargs, self.master)
    #     wlc_dict = {}
    #     for tp in self.iter_list:
    #         self.group = get_group(self.master, tp, self.kwargs)
    #         wlc_dict = {}
    #         p_total = 0  # portfolio total
    #
    #         for i, g in enumerate(self.group):
    #             l_group = get_group(self.master, tp, self.kwargs, i)  # lower group
    #             g_total = 0
    #             l_g_l = []  # lower group list
    #             for p in l_group:
    #                 p_data = get_correct_p_data(
    #                     self.kwargs, self.master, self.baseline_type, p, tp
    #                 )
    #                 wlc = p_data["Total Forecast"]
    #                 if isinstance(wlc, (float, int)) and wlc is not None and wlc != 0:
    #                     if wlc > 50000:
    #                         logger.info(
    #                             tp
    #                             + ", "
    #                             + str(p)
    #                             + " is £"
    #                             + str(round(wlc))
    #                             + " please check this is correct. For now analysis_engine has recorded it as £0"
    #                         )
    #                     # wlc_dict[p] = wlc
    #                 if wlc == 0:
    #                     logger.info(
    #                         tp
    #                         + ", "
    #                         + str(p)
    #                         + " wlc is currently £"
    #                         + str(wlc)
    #                         + " note this is key information that should be provided by the project"
    #                     )
    #                     # wlc_dict[p] = wlc
    #                 if wlc is None:
    #                     logger.info(
    #                         tp
    #                         + ", "
    #                         + str(p)
    #                         + " wlc is currently None note this is key information that should be provided by the project"
    #                     )
    #                     wlc = 0
    #
    #                 l_g_l.append((wlc, p))
    #                 g_total += wlc
    #
    #             wlc_dict[g] = list(reversed(sorted(l_g_l)))
    #             p_total += g_total
    #
    #         wlc_dict["total"] = p_total
    #         wlc_dict[tp] = wlc_dict
    #
    #     self.wlc_dict = wlc_dict
    #
    # def calculate_wlc_change(self) -> None:
    #     wlc_change_dict = {}
    #     for i, tp in enumerate(self.wlc_dict.keys()):
    #         p_wlc_change_dict = {}
    #         for p in self.wlc_dict[tp].keys():
    #             wlc_one = self.wlc_dict[tp][p]
    #             try:
    #                 wlc_two = self.wlc_dict[self.iter_list[i + 1]][p]
    #                 try:
    #                     percentage_change = int(((wlc_one - wlc_two) / wlc_one) * 100)
    #                     p_wlc_change_dict[p] = percentage_change
    #                 except ZeroDivisionError:
    #                     logger.info(
    #                         "As "
    #                         + str(p)
    #                         + " has no wlc total figure for "
    #                         + tp
    #                         + " change has been calculated as zero"
    #                     )
    #             except IndexError:  # handles NoneTypes.
    #                 pass
    #
    #         wlc_change_dict[tp] = p_wlc_change_dict
    #
    #     self.wlc_change = wlc_change_dict
