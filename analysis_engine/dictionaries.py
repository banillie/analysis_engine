


RAG_RANKING_DICT_NUMBER = {
    6: 'Green',  # in case combo of green and improving.
    5: 'Green',
    4: 'Amber/Green',
    3: 'Amber',
    2: 'Amber/Red',
    1: 'Red',
    0: None
}
RAG_RANKING_DICT_COLOUR = {
    'Green': 5,
    'Amber': 3,
    'Red': 1,
    None: 0,
}

BC_STAGE_DICT = {
    "Strategic Outline Case": "SOBC",
    "SOBC": "SOBC",
    "pre-Strategic Outline Case": "pre-SOBC",
    "pre-SOBC": "pre-SOBC",
    "Outline Business Case": "OBC",
    "OBC": "OBC",
    "Full Business Case": "FBC",
    "FBC": "FBC",
    # older returns that require cleaning
    "Pre - SOBC": "pre-SOBC",
    "Pre Strategic Outline Business Case": "pre_SOBC",
    None: None,
    "Other": "Other",
    "Other ": "Other",
    "To be confirmed": None,
    "To be confirmed ": None,
    "Ongoing Board papers": None,
}

DCA_KEYS = {
    "sro": "Departmental DCA",
    "finance": "SRO Finance confidence",
    "benefits": "SRO Benefits RAG",
    "schedule": "SRO Schedule Confidence",
    "resource": "Overall Resource DCA - Now",
}

STANDARDISE_DCA_KEYS = {
    'cdg': 'Overall Delivery Confidence',
    'top_250': None,
    'ipdc': 'Departmental DCA',
}

FONT_TYPE = ["sans serif", "Ariel"]


BEN_TYPE_KEY_LIST = [
    (
        "Pre-profile BEN Forecast Gov Cashable",
        "Pre-profile BEN Forecast Gov Non-Cashable",
        "Pre-profile BEN Forecast - Economic (inc Private Partner)",
        "Pre-profile BEN Forecast - Disbenefit UK Economic",
    ),
    (
        "Total BEN Forecast - Gov. Cashable",
        "Total BEN Forecast - Gov. Non-Cashable",
        "Total BEN Forecast - Economic (inc Private Partner)",
        "Total BEN Forecast - Disbenefit UK Economic",
    ),
    (
        "Unprofiled Remainder BEN Forecast - Gov. Cashable",
        "Unprofiled Remainder BEN Forecast - Gov. Non-Cashable",
        "Unprofiled Remainder BEN Forecast - Economic (inc Private Partner)",
        "Unprofiled Remainder BEN Forecast - Disbenefit UK Economic",
    ),
]

YEAR_LIST = [
    # "16-17",
    # "17-18",
    # "18-19",
    # "19-20",
    # "20-21",
    "21-22",
    "22-23",
    "23-24",
    "24-25",
    "25-26",
    "26-27",
    "27-28",
    "28-29",
    "29-30",
    "30-31",
    "31-32",
    "32-33",
    "33-34",
    "34-35",
    "35-36",
    "36-37",
    "37-38",
    "38-39",
    "39-40",
]

COST_KEY_LIST = [
    " RDEL Forecast Total",
    " CDEL Forecast one off new costs",
    " Forecast Non-Gov",
]

STANDARDISE_COST_KEYS = {
    'spent': {'cdg': 'Total Costs Spent'},
    'remaining': {'cdg': 'Total Costs Remaining'},
    'total': {'cdg': 'Total Costs'},
    'income_achieved': {'cdg': 'Total Income Achieved'},
    'income_remaining': {'cdg': 'Total Income Remaining'},
    'income_total': {'cdg': 'Total Income'}
}


def convert_rag_text(dca_rating: str) -> str:
    """Converts RAG name into a acronym"""

    if dca_rating == "Green":
        return "G"
    elif dca_rating == "Amber/Green":
        return "A/G"
    elif dca_rating == "Amber":
        return "A"
    elif dca_rating == "Amber/Red":
        return "A/R"
    elif dca_rating == "Red":
        return "R"
    else:
        return ""