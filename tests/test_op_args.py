# should include some settings that are wrong to test error messaging
REPORTING_TYPE = "cdg"
chart = "show"

q_koi = {
    'test_name': 'query_koi',
    'subparser_name': 'query',
    'koi': 'Brief Description'
}
q_koi_two_keys = {
    'test_name': 'query_koi_two_keys',
    'subparser_name': 'query',
    'koi': ['Brief Description', 'Last Business Case (BC) achieved']
}
query_koi_quarters = {
    'test_name': 'query_koi_quarters',
    'subparser_name': 'query',
    'koi': ['Brief Description', 'Last Business Case (BC) achieved'],
    "quarter": ["Q1 22/23", "Q4 21/22", "Q3 21/22"],
}
query_koi_milestones = {
    'test_name': 'query_koi_milestones',
    'subparser_name': 'query',
    'koi': "FBC CDG Approval",
    "quarter": ["Q1 22/23", "Q4 21/22", "Q3 21/22"],
}
q_koi_fn = {
    'test_name': 'query_koi_fn',
    'subparser_name': 'query',
    'koi_fn': 'test_query_keys',
    "quarter": ["Q1 22/23", "Q4 21/22", "Q3 21/22"],
}

QUERY_ARGS = [
    q_koi_fn,
    query_koi_milestones,
    query_koi_quarters,
    q_koi_two_keys,
    q_koi
]

ms_standard = {
    "test_name": "ms_standard",
    "subparser_name": "milestones",
    "chart": chart,
}
ms_groups = {
    "test_name": "ms_groups",
    "subparser_name": "milestones",
    "group": ["SCS"],
    "chart": chart,
}
ms_dates = {
    "test_name": "ms_dates",
    "subparser_name": "milestones",
    "chart": chart,
    "dates": ["1/2/2022", "1/2/2023"],
}
ms_blue_line_config = {
    "test_name": "ms_bl_config",
    "subparser_name": "milestones",
    "blue_line": "config_date",
    "chart": chart,
}
ms_blue_line_today = {
    "test_name": "ms_bl_today",
    "subparser_name": "milestones",
    "blue_line": "today",
    "chart": chart,
}
ms_koi = {
    "test_name": "ms_koi",
    "subparser_name": "milestones",
    "chart": chart,
    "koi": "FBC CDG Approval",
}
ms_koi_fn = {
    "test_name": "ms_koi_fn",
    "subparser_name": "milestones",
    "chart": chart,
    "koi_fn": "milestone_keys",
}
ms_quarters = {
    "test_name": "ms_quarters",
    "subparser_name": "milestones",
    "chart": chart,
    "quarter": ["Q4 21/22", "Q3 21/22", "Q2 21/22"],
}
ms_groups = {
    "test_name": "ms_groups",
    "subparser_name": "milestones",
    "group": ["SCS"],
    "chart": chart,
}
ms_stages = {
    "test_name": "ms_stages",
    "subparser_name": "milestones",
    "stage": [
        # 'Outline Business Case',
        # 'Full Business Case',
        "Ongoing Board papers"
    ],
    "chart": chart,
}

MILESTONES_OP_ARGS = [
    ms_stages,
    ms_groups,
    ms_quarters,
    ms_koi_fn,
    ms_koi,
    ms_blue_line_today,
    ms_blue_line_config,
    ms_dates,
    ms_groups,
    ms_standard,
]

dlion_standard = {
    "test_name": "dlion_standard",
    "subparser_name": "dandelion",
    "chart": chart,
}
dlion_groups = {
    "test_name": "dlion_groups",
    "subparser_name": "dandelion",
    "group": ["SCS", "CFPD"],
    "chart": chart,
}
dlion_stages = {
    "test_name": "dlion_stages",
    "subparser_name": "dandelion",
    "stage": [
        # 'pre-Strategic Outline Case',
        # 'Strategic Outline Case',
        # 'Outline Business Case',
        "Full Business Case",
        "Ongoing Board papers",
    ],
    "chart": chart,
}
dlion_quarter = {
    "test_name": "dlion_quarter",
    "subparser_name": "dandelion",
    "quarter": ["Q2 21/22"],
    "group": ["SCS", "CFPD", "GF"],
    "chart": chart,
}
dlion_angles = {
    "test_name": "dlion_angles",
    "subparser_name": "dandelion",
    "angles": [280, 360, 80],
    "chart": chart,
}
dlion_benefits = {
    "test_name": "dlion_benefits",
    "subparser_name": "dandelion",
    "type": "benefits",
    "chart": chart,
}
dlion_income = {
    "test_name": "dlion_income",
    "angles": [280, 360, 80],
    "subparser_name": "dandelion",
    "type": "income",
    "chart": chart,
}
# this will crash if project not in quarter data master
dlion_cli_group = {
    "test_name": "dlion_groups",
    "subparser_name": "dandelion",
    "group": ["WIT retail project", "Mayfield", "MSG"],
    "chart": chart,
}

DANDELION_OP_ARGS_DICT = [
    # # dlion_cli_group,  # Failing.
    dlion_income,
    dlion_benefits,
    dlion_angles,
    dlion_quarter,
    dlion_stages,
    dlion_groups,
    dlion_standard,
]

sd_standard = {
    "test_name": "sd_standard",
    "subparser_name": "speed_dials",
    "rag_number": "5",  # NOT REQUIRED FOR DCA ANALYSIS
    "quarter": "standard",
}
sd_quarters = {
    "test_name": "sd_quarters",
    "subparser_name": "speed_dials",
    "rag_number": "5",
    "quarter": ["Q2 21/22", "Q4 21/22"],
}
sd_stage = {
    "test_name": "sd_stage",
    "subparser_name": "speed_dials",
    "rag_number": "5",
    "quarter": "standard",
    "stage": [
        "OBC",
        "FBC",
    ],
}
sd_groups = {
    "test_name": "sd_groups",
    "subparser_name": "speed_dials",
    "rag_number": "5",
    "quarter": "standard",
    "group": ["SCS", "CFPD"],
}

SPEED_DIAL_AND_DCA_OP_ARGS = [
    sd_groups,
    sd_stage,
    sd_quarters,
    sd_standard
]


# 'rag_number': '5',
# 'order_by': 'cost',
# 'none_handle': 'none',
