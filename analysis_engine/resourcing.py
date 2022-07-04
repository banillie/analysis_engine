from analysis_engine.segmentation import get_iter_list, get_group, get_correct_p_data
from analysis_engine.cleaning import convert_none_types


class ResourceData:
    def __init__(self, master, **kwargs):
        self.master = master
        self.baseline_type = "ipdc_costs"
        self.kwargs = kwargs
        self.start_group = []
        self.group = []
        self.iter_list = []
        self.get_resource_totals()
        # self.ps_resource = 0
        # self.contractor_resource = 0
        # self.total_resource = 0

    def get_resource_totals(self) -> None:
        self.iter_list = get_iter_list(self.kwargs, self.master)
        for tp in self.iter_list:
            self.group = get_group(self.master, tp, self.kwargs)
            public_sector_resource = []
            c_resource = []
            t_resource = []
            fp_resource = []
            for project_name in self.group:
                p_data = get_correct_p_data(
                    self.kwargs, self.master, self.baseline_type, project_name, tp
                )
                if p_data is None:
                    break
                else:
                    ps = convert_none_types(p_data["DfTc Public Sector Employees"])
                    public_sector_resource.append(ps)
                    c = convert_none_types(p_data["DfTc External Contractors"])
                    c_resource.append(c)
                    t = convert_none_types(p_data["DfTc Project Team Total"])
                    t_resource.append(t)
                    fp = convert_none_types(p_data["DfTc Funded Posts"])
                    fp_resource.append(fp)

        self.ps_resource = sum(public_sector_resource)
        self.contractor_resource = sum(c_resource)
        self.total_resource = sum(t_resource)
        self.funded = sum(fp_resource)
