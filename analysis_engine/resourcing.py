from analysis_engine.dictionaries import RESOURCE_KEYS, RESOURCE_KEYS_OLD
from analysis_engine.segmentation import get_iter_list, get_group, get_correct_p_data
from analysis_engine.cleaning import convert_none_types
from analysis_engine.error_msgs import logger, resourcing_keys


class ResourceData:
    def __init__(self, master, **kwargs):
        self.master = master
        self.quarters = self.master["quarter_list"]
        self.kwargs = kwargs
        self.totals = {}
        self.get_resource_totals()

    def get_resource_totals(self) -> None:
        totals_dict = {}
        for tp in self.quarters:
            public_sector_resource = []
            c_resource = []
            t_resource = []
            fp_resource = []
            group = get_group(self.master, tp, **self.kwargs)
            for project_name in group:
                p_data = get_correct_p_data(self.master, project_name, tp)

                try:
                    ps = convert_none_types(p_data[RESOURCE_KEYS["ps_resource"]])
                except KeyError:
                    ps = convert_none_types(p_data[RESOURCE_KEYS_OLD["ps_resource"]])
                if type(ps) == str:
                    resourcing_keys(project_name, RESOURCE_KEYS_OLD['ps_resource'])
                    public_sector_resource.append(0)
                else:
                    public_sector_resource.append(ps)

                try:
                    c = convert_none_types(p_data[RESOURCE_KEYS["contractor_resource"]])
                except KeyError:
                    c = convert_none_types(
                        p_data[RESOURCE_KEYS_OLD["contractor_resource"]]
                    )
                if type(c) == str:
                    resourcing_keys(project_name, RESOURCE_KEYS_OLD["contractor_resource"])
                    c_resource.append(0)
                else:
                    c_resource.append(c)

                try:
                    t = convert_none_types(p_data[RESOURCE_KEYS["total_resource"]])
                except KeyError:
                    t = convert_none_types(p_data[RESOURCE_KEYS_OLD["total_resource"]])
                if type(t) == str:
                    resourcing_keys(project_name, RESOURCE_KEYS_OLD["total_resource"])
                    t_resource.append(0)
                else:
                    t_resource.append(t)

                try:
                    fp = convert_none_types(p_data[RESOURCE_KEYS["funded_resource"]])
                except KeyError:
                    fp = convert_none_types(
                        p_data[RESOURCE_KEYS_OLD["funded_resource"]]
                    )
                if type(t) == str:
                    resourcing_keys(project_name, RESOURCE_KEYS_OLD["funded_resource"])
                    fp_resource.append(0)
                else:
                    fp_resource.append(fp)

            totals_dict[tp] = {
                "ps_resource": sum(public_sector_resource),
                "contractor_resource": sum(c_resource),
                "total_resource": sum(t_resource),
                "funded_resource": sum(fp_resource),
            }

        self.totals = totals_dict
