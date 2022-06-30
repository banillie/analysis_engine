REPORTING_TYPE = "cdg"

chart = False

standard = {
    "quarter": ["Q4 21/22"],
    "group": ["SCS", "CFPD", "GF"],
    'chart': chart,
}

two_groups = {"quarter": ["Q4 21/22"], "group": ["SCS", "CFPD"], 'chart': chart}

stages = {
    "quarter": ["Q4 21/22"],
    "stage": [
        "pre-Strategic Outline Case",
        "Strategic Outline Case",
        "Outline Business Case",
        "Full Business Case",
        "Ongoing Board papers",
        ],
    'chart': chart
}

change_quarter = {
    "quarter": ["Q2 19/22"],
    "group": ["SCS", "CFPD", "GF"],
    'chart': True,
}


OP_ARGS_DICT = [
    change_quarter,
    stages,
    standard,
    two_groups,
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
