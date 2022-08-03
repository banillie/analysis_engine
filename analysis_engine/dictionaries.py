RAG_RANKING_DICT_NUMBER = {  # for dandelion
    6: "Green",  # in case combo of green and improving.
    5: "Green",
    4: "Amber/Green",
    3: "Amber",
    2: "Amber/Red",
    1: "Red",
    0: None,
}
RAG_RANKING_DICT_COLOUR = {  # for dandelion
    "Green": 5,
    "Amber": 3,
    "Red": 1,
    None: 0,
}

# BC stage terms are consistent across reports
BC_STAGE_DICT_ABB_TO_FULL = {
    "SOBC": "Strategic Outline Case",
    "pre-SOBC": "pre-Strategic Outline Case",
    "OBC": "Outline Business Case",
    "FBC": "Full Business Case",
    "Ongoing Board papers": "Ongoing Board papers",
}

BC_STAGE_DICT_FULL_TO_ABB = {
    "Strategic Outline Case": "SOBC",
    "pre-Strategic Outline Case": "pre-SOBC",
    "Outline Business Case": "OBC",
    "Full Business Case": "FBC",
    "Ongoing Board papers": "OBPs",
}

DCA_KEYS = {
    "cdg": {
        "sro": "Overall Delivery Confidence",
        "finance": "Costs Confidence",
        "benefits": "Benefits Confidence",
        "schedule": "Schedule Confidence",
    },
    "ipdc": {
        "sro": "Departmental DCA",
        "finance": "SRO Finance confidence",
        "benefits": "SRO Benefits RAG",
        "schedule": "SRO Schedule Confidence",
    },
    "ipa": {
        "ipa": "GMPP - IPA DCA",
    },
    "resource": {"resource": "Overall Resource DCA - Now"},
}

PROJECT_INFO_KEYS = {
    "cdg": {
        "group": "Directorate",
    },
    "ipdc": {
        "group": "Group",
    },
}

# rationalise with RAG_RANKING_DICT_COLOUR
DCA_RATING_SCORES = {
    "Green": 5,
    "Amber/Green": 4,
    "Amber": 3,
    "Amber/Red": 2,
    "Red": 1,
    None: None,
}

# STANDARDISE_DCA_KEYS = {
#     "cdg": "Overall Delivery Confidence",
#     "top_250": None,
#     "ipdc": "Departmental DCA",
# }

FONT_TYPE = ["sans serif", "Ariel"]

STANDARDISE_COST_KEYS = {
    "cdg": {
        "spent": "Total Costs Spent",
        "remaining": "Total Costs Remaining",
        "total": "Total Costs",
        "income_achieved": "Total Income Achieved",
        "income_remaining": "Total Income Remaining",
        "income_total": "Total Income",
    },
    "ipdc": {
        "spent": "Spent Costs",
        "remaining": None,
        "total": "Total Forecast",
        "income_achieved": None,
        "income_remaining": None,
        "income_total": "Total Forecast - Income both Revenue and Capital",
    },
}

STANDARDISE_BEN_KEYS = {
    "delivered": {"cdg": "Benefits delivered"},
    "remaining": {"cdg": "Benefits to be delivered"},
    "total": {"cdg": "Total Benefits"},
}

rag_txt_list = ["A/G", "A/R", "R", "G", "A"]  # cdg dashboards

conf_list = [
    "Costs Confidence",
    "Schedule Confidence",
    "Benefits Confidence",
]  # cdg dashboards
risk_list = [
    "Benefits",
    "Capability",
    "Cost",
    "Objectives",
    "Purpose",
    "Schedule",
    "Sponsorship",
    "Stakeholders",
]  # cdg dashboard


# ONLY USED FOR CDG DASHBOARDS AT MOMENT
DATA_KEY_DICT = {
    "IPDC approval point": "Last Business Case (BC) achieved",
    "Total Forecast": "Total Costs",
    "Departmental DCA": "Overall Delivery Confidence",
}

# Used in dashboards
CONVERT_RAG = {
    "Green": "G",
    "Amber/Green": "A/G",
    "Amber": "A",
    "Amber/Red": "A/R",
    "Red": "R",
    None: None,
}

NEXT_STAGE_DICT = {
    "pre-SOBC": "SOBC - IPDC Approval",
    "SOBC": "OBC - IPDC Approval",
    "OBC": "FBC - IPDC Approval",
    "FBC": "Project End Date",
    "Other": None,
}

DANDELION_KEYS = {
    "forward_look": "SRO Forward Look Assessment",
}

FWD_LOOK_DICT = {
    "Worsening": 1,
    "No Change Expected": 2,
    "Improving": 3,
    None: "",
}

DASHBOARD_KEYS = {
    "BC_STAGE": "IPDC approval point",
    # "CONTINGENCY": "Overall contingency (£m)",   # not used
    # "OB": "Overall figure for Optimism Bias (£m)",  # not used
}

RESOURCE_KEYS_OLD = {
    "ps_resource": "DfTc Public Sector Employees",
    "contractor_resource": "DfTc External Contractors",
    "total_resource": "DfTc Project Team Total",
    "funded_resource": "DfTc Funded Posts",
}

RESOURCE_KEYS = {
    "ps_resource": "No of DfTc FTEs working on Project",
    "contractor_resource": "Number of External Contractors (FTEs)",
    "total_resource": "Total (FTEs)",
    "funded_resource": "Total Number of Funded Posts (FTEs)",
}

SCHEDULE_DASHBOARD_KEYS = [
    "Start of Construction/build",
    "Start of Operation",
    "Full Operations",
    "Project End Date",
]

DASHBOARD_RESOURCE_KEYS = [
    "No of DfTc FTEs working on Project",
    "Number of External Contractors (FTEs)",
    "Total (FTEs)",
    "Total Number of Funded Posts (FTEs)",
    "Resource Gap",
    "DfTc Resource Gap Criticality (RAG rating)",
]

RISK_LIST = [
    "Brief Risk Description ",
    "BRD Risk Category",
    "BRD Primary Risk to",
    "BRD Internal Control",
    "BRD Mitigation - Actions taken (brief description)",
    "BRD Residual Impact",
    "BRD Residual Likelihood",
    "Severity Score Risk Category",
    "BRD Has this Risk turned into an Issue?",
]

PORTFOLIO_RISK_LIST = [
    "Portfolio Risk Impact Description",
    "Portfolio Risk Mitigation",
    "Portfolio Risk Likelihood",
    "Portfolio Risk Impact Assessment",
    "Severity Score Risk Category",
]

RISK_SCORES = {
    "Very Low": 0,
    "Low": 1,
    "Medium": 2,
    "High": 3,
    "Very High": 4,
    "N/A": None,
    None: None,
}

PORTFOLIO_RISK_SCORES = {
    "N/A": None,
    "Unlikely": 1,
    None: None,
    "Very Unlikely": 0,
    "Likely": 3,
    "Possible": 2,
    "Very Likely": 4,
}