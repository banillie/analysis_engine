# should include some settings that are wrong to test error messaging

REPORTING_TYPE = "cdg"

dlion_chart = False

dlion_standard = {
    "chart": dlion_chart,
}

dlion_groups = {"group": ["SCS", "CFPD"], "chart": dlion_chart}

dlion_stages = {
    "stage": [
        "pre-Strategic Outline Case",
        "Strategic Outline Case",
        "Outline Business Case",
        "Full Business Case",
        "Ongoing Board papers",
    ],
    "chart": dlion_chart,
}

dlion_quarter = {
    "quarter": ["Q2 21/22"],
    "group": ["SCS", "CFPD", "GF"],
    "chart": dlion_chart,
}

dlion_angles = {
    'angles': [300, 360, 60],
    'chart': dlion_chart
}

dlion_benefits = {
    'type': 'benefits',
    'chart': dlion_chart,
}

dlion_income = {
    'type': 'income',
    'chart': dlion_chart,
}

DANDELION_OP_ARGS_DICT = [
    dlion_income,
    dlion_benefits,
    dlion_angles,
    dlion_quarter,
    dlion_stages,
    dlion_groups,
    dlion_standard,
]

sd_standard = {
    'rag_number': '5',  # NOT REQUIRED FOR DCA ANALYSIS
    'quarter': 'standard',
}

sd_quarters = {
    'rag_number': '5',
    'quarter': ['Q2 21/22', 'Q4 21/22'],
}

sd_stage = {
    'rag_number': '5',
    'quarter': 'standard',
    "stage": [
        "OBC",
        "FBC",
    ]
}

sd_groups = {
    'rag_number': '5',
    'quarter': 'standard',
    "group": ["SCS", "CFPD"],
}

SPEED_DIAL_AND_DCA_OP_ARGS = [
    sd_groups,
    sd_stage,
    sd_quarters,
    sd_standard
]

# "group": ["SCS", "GF"],
# "chart": True,
# "type": "income",
# "blue_line": "CDG",
# "dates": ["1/10/2021", "1/10/2022"],
# "fig_size": "half_horizontal",
# "rag_number": "5",
# "order_by": "cost",
# "angles": [300, 360, 60],
# "none_handle": "none",
# "quarter": ["standard"],
