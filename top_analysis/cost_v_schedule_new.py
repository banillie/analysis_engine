"""
New code for cost v schedule matrix
"""

from data_mgmt.data import (
    Master,
    get_project_information,
    get_master_data, CostData,
    MilestoneData
)


m = Master(get_master_data(), get_project_information())
costs = CostData(m, m.current_projects)
milestones = MilestoneData(m, m.current_projects)
