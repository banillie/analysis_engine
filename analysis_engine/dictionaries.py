


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