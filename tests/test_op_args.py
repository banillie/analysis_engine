REPORTING_TYPE = "cdg"

chart = True

## should include some settings that are wrong to test error messaging

standard = {
    "chart": chart,
}

two_groups = {"group": ["SCS", "CFPD"], "chart": chart}

stages = {
    "stage": [
        "pre-Strategic Outline Case",
        "Strategic Outline Case",
        "Outline Business Case",
        "Full Business Case",
        "Ongoing Board papers",
    ],
    "chart": chart,
}

change_quarter = {
    "quarter": ["Q2 21/22"],
    "group": ["SCS", "CFPD", "GF"],
    "chart": chart,
}

change_angles = {
    'angles': [300, 360, 60],
    'chart': chart
}

benefits = {
    'type': 'benefits',
    'chart': chart,
}

OP_ARGS_DICT = [
    benefits,
    # change_angles,
    # change_quarter,
    # stages,
    # two_groups,
    # standard,
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
