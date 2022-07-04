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

SPEED_DIAL_OP_ARGS = {
    'rag_number': '5',
    'quarter': 'standard',
}

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
