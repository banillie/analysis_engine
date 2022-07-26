from analysis_engine.segmentation import (
    get_iter_list,
    get_group,
    get_correct_p_data,
    calculate_profiled,
)
from analysis_engine.cleaning import convert_none_types
from analysis_engine.dictionaries import STANDARDISE_BEN_KEYS


class BenefitsData:
    """
    Note currently in use for ipdc reporting. requires refactor.
    """

    def __init__(self, master, **kwargs):
        self.master = master
        self.baseline_type = "ipdc_benefits"
        self.kwargs = kwargs
        self.report = kwargs["report"]
        self.b_totals = {}
        self.get_ben_totals()

    def get_ben_totals(self) -> None:
        """
        Returns lists containing benefit totals
        """
        lower_dict = {}
        for tp in self.kwargs["quarter"]:
            delivered = 0
            remaining = 0
            total = 0
            group = get_group(self.master, tp, **self.kwargs)
            for project_name in group:
                p_data = get_correct_p_data(self.master, project_name, tp)
                delivered += convert_none_types(
                    p_data[STANDARDISE_BEN_KEYS["delivered"][self.report]]
                )
                remaining += convert_none_types(
                    p_data[STANDARDISE_BEN_KEYS["remaining"][self.report]]
                )
                total += convert_none_types(
                    p_data[STANDARDISE_BEN_KEYS["total"][self.report]]
                )

            lower_dict[tp] = {
                "delivered": delivered,
                "prof": remaining,
                "total": total,
            }

        # else:
        #     for tp in self.iter_list:
        #         delivered = 0
        #         profiled = 0
        #         unprofiled = 0
        #         cash_dev = 0
        #         uncash_dev = 0
        #         economic_dev = 0
        #         disben_dev = 0
        #         cash_profiled = 0
        #         uncash_profiled = 0
        #         economic_profiled = 0
        #         disben_profiled = 0
        #         cash_unprofiled = 0
        #         uncash_unprofiled = 0
        #         economic_unprofiled = 0
        #         disben_unprofiled = 0
        #         self.group = get_group(self.master, tp, self.kwargs)
        #         for x, key in enumerate(BEN_TYPE_KEY_LIST):
        #             # group_total = 0
        #             for p in self.group:
        #                 p_data = get_correct_p_data(
        #                     self.kwargs, self.master, self.baseline_type, p, tp
        #                 )
        #                 if p_data is None:
        #                     continue
        #                 try:
        #                     cash = round(p_data[key[0]])
        #                     if cash is None:
        #                         cash = 0
        #                     uncash = round(p_data[key[1]])
        #                     if uncash is None:
        #                         uncash = 0
        #                     economic = round(p_data[key[2]])
        #                     if economic is None:
        #                         economic = 0
        #                     disben = round(p_data[key[3]])
        #                     if disben is None:
        #                         disben = 0
        #                     total = round(cash + uncash + economic + disben)
        #                     # group_total += total
        #                 except TypeError:  # handle None types, which are present if project not reporting last quarter.
        #                     cash = 0
        #                     uncash = 0
        #                     economic = 0
        #                     disben = 0
        #                     total = 0
        #                     # group_total += total
        #
        #                 if self.iter_list.index(tp) == 0:  # current quarter
        #                     if x == 0:  # spent
        #                         cash_dev += cash
        #                         uncash_dev += uncash
        #                         economic_dev += economic
        #                         disben_dev += disben
        #                     if x == 1:  # profiled
        #                         cash_profiled += cash
        #                         uncash_profiled += uncash
        #                         economic_profiled += economic
        #                         disben_profiled += disben
        #                     if x == 2:  # unprofiled
        #                         cash_unprofiled += cash
        #                         uncash_unprofiled += uncash
        #                         economic_unprofiled += economic
        #                         disben_unprofiled += disben
        #
        #                 if x == 0:  # spent
        #                     delivered += total
        #                 if x == 1:  # profiled
        #                     profiled += total
        #                 if x == 2:  # unprofiled
        #                     unprofiled += total
        #
        #         cat_spent = [cash_dev, uncash_dev, economic_dev, disben_dev]
        #         cat_profiled = [
        #             cash_profiled,
        #             uncash_profiled,
        #             economic_profiled,
        #             disben_profiled,
        #         ]
        #         cat_unprofiled = [
        #             cash_unprofiled,
        #             uncash_unprofiled,
        #             economic_unprofiled,
        #             disben_unprofiled,
        #         ]
        #         cat_profiled = calculate_profiled(
        #             cat_profiled, cat_spent, cat_unprofiled
        #         )
        #         adj_profiled = calculate_profiled(profiled, delivered, unprofiled)
        #         lower_dict[tp] = {
        #             "cat_spent": cat_spent,
        #             "cat_prof": cat_profiled,
        #             "cat_unprof": cat_unprofiled,
        #             "delivered": delivered,
        #             "prof": adj_profiled,
        #             "unprof": unprofiled,
        #             "total": profiled,
        #         }

        self.b_totals = lower_dict
