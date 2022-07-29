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

# BEN_TYPE_KEY_LIST = [
#     (
#         "Pre-profile BEN Forecast Gov Cashable",
#         "Pre-profile BEN Forecast Gov Non-Cashable",
#         "Pre-profile BEN Forecast - Economic (inc Private Partner)",
#         "Pre-profile BEN Forecast - Disbenefit UK Economic",
#     ),
#     (
#         "Total BEN Forecast - Gov. Cashable",
#         "Total BEN Forecast - Gov. Non-Cashable",
#         "Total BEN Forecast - Economic (inc Private Partner)",
#         "Total BEN Forecast - Disbenefit UK Economic",
#     ),
#     (
#         "Unprofiled Remainder BEN Forecast - Gov. Cashable",
#         "Unprofiled Remainder BEN Forecast - Gov. Non-Cashable",
#         "Unprofiled Remainder BEN Forecast - Economic (inc Private Partner)",
#         "Unprofiled Remainder BEN Forecast - Disbenefit UK Economic",
#     ),
# ]
#
# YEAR_LIST = [
#     # "16-17",
#     # "17-18",
#     # "18-19",
#     # "19-20",
#     # "20-21",
#     "21-22",
#     "22-23",
#     "23-24",
#     "24-25",
#     "25-26",
#     "26-27",
#     "27-28",
#     "28-29",
#     "29-30",
#     "30-31",
#     "31-32",
#     "32-33",
#     "33-34",
#     "34-35",
#     "35-36",
#     "36-37",
#     "37-38",
#     "38-39",
#     "39-40",
# ]

# COST_KEY_LIST = [
#     " RDEL Forecast Total",
#     " CDEL Forecast one off new costs",
#     " Forecast Non-Gov",
# ]

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
        "spent": None,
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

DASHBOARD_KEYS = {
    "BC_STAGE": "IPDC approval point",
    "CONTINGENCY": "Overall contingency (£m)",
    "OB": "Overall figure for Optimism Bias (£m)",
}
