from data_mgmt.data import (
    Master,
    get_master_data,
    get_project_information,
    MilestoneData,
    Projects,
    milestone_chart,
    save_graph,
)

PARLIAMENT = [
    "Bill",
    "bill",
    "hybrid",
    "Hybrid",
    "reading",
    "royal",
    "Royal",
    "assent",
    "Assent",
    "legislation",
    "Legislation",
    "Passed",
    "NAO",
    "nao",
    "PAC",
    "pac",
]
CONSTRUCTION = [
    "Start of Construction/build",
    "Complete",
    "complete",
    "Tender",
    "tender",
]
OPERATIONS = [
    "Full Operations",
    "Start of Operation",
    "operational",
    "Operational",
    "operations",
    "Operations",
    "operation",
    "Operation",
]
OTHER_GOV = ["TAP", "MPRG", "Cabinet Office", " CO ", "HMT"]
CONSULTATIONS = [
    "Consultation",
    "consultation",
    "Preferred",
    "preferred",
    "Route",
    "route",
    "Announcement",
    "announcement",
    "Statutory",
    "statutory",
    "PRA",
]
PLANNING = [
    "DCO",
    "dco",
    "Planning",
    "planning",
    "consent",
    "Consent",
    "Pre-PIN",
    "Pre-OJEU",
    "Initiation",
    "initiation",
]
IPDC = ["IPDC", "BICC"]
HE_SPECIFIC = [
    "Start of Construction/build",
    "DCO",
    "dco",
    "PRA",
    "Preferred",
    "preferred",
    "Route",
    "route",
    "Annoucement",
    "announcement",
    "submission",
    "PVR" "Submission",
]

master = Master(get_master_data(), get_project_information())
milestones = MilestoneData(master, master.current_projects)
milestones.filter_chart_info(milestone_type='Delivery', start_date="1/12/2020", end_date="1/6/2021")
f = milestone_chart(milestones, title="Planning")
# save_graph(f, "non zero solution 2")
