import datetime
from dateutil.relativedelta import relativedelta


def date_past(time_period: int):
    # helper function for setting test dates
    month_ago = datetime.datetime.today() - relativedelta(months=+time_period)
    return month_ago.strftime('%d-%m-%Y')


def date_future(time_period: int):
    # helper function for setting test dates
    three_months = datetime.datetime.today() + relativedelta(months=+time_period)
    return three_months.strftime('%d-%m-%Y')


# should include some settings that are wrong to test error messaging
REPORTING_TYPE = "ipdc"
chart = "show"

cost_standard = {
    "test_name": "cost_standard",
    "subparser_name": "costs",
    "chart": chart,
}
cost_quarters = {
    "test_name": "cost_quarters",
    "subparser_name": "costs",
    "chart": chart,
    "quarter": ["Q1 22/23", "Q4 21/22", "Q1 21/22"],
}
cost_groups = {
    "test_name": "cost_groups",
    "subparser_name": "costs",
    "group": ["LTC"],
    "quarter": ["Q1 22/23", "Q4 21/22", "Q1 21/22"],
    "chart": chart,
}
cost_baseline = {
    "test_name": "cost_baseline",
    "subparser_name": "costs",
    "chart": chart,
    "baseline": [],
}
cost_baseline_remove = {
    "test_name": "cost_baseline_remove",
    "subparser_name": "costs",
    "chart": chart,
    "baseline": [],
    "remove": ["HS2 Ph 1"]
}
cost_baseline_quarter = {
    "test_name": "cost_baseline_qrt",
    "subparser_name": "costs",
    "chart": chart,
    "baseline": [],
    "quarter" : ["Q4 21/22"]
}

if REPORTING_TYPE == 'ipdc':
    COST_OP_ARGS = [
        cost_baseline_quarter,
        cost_baseline_remove,
        cost_baseline,
        cost_groups,
        cost_standard,
        cost_quarters,
    ]


q_koi_cdg = {
    'test_name': 'query_koi',
    'subparser_name': 'query',
    'koi': 'Brief Description'
}
q_koi_ipdc = {
    'test_name': 'query_koi',
    'subparser_name': 'query',
    'koi': 'Departmental DCA'
}
q_koi_two_keys_ipdc = {
    'test_name': 'query_koi_two_keys',
    'subparser_name': 'query',
    'koi': ['Snapshot Date', 'IO4 - Monetised?']
}
q_koi_two_keys_cdg = {
    'test_name': 'query_koi_two_keys',
    'subparser_name': 'query',
    'koi': ['Brief Description', 'Last Business Case (BC) achieved']
}
query_koi_quarters_ipdc = {
    'test_name': 'query_koi_quarters',
    'subparser_name': 'query',
    'koi': ['Departmental DCA'],
    "quarter": ["Q4 21/22", "Q3 21/22", "Q2 21/22"],
}
query_koi_quarters_cdg = {
    'test_name': 'query_koi_quarters',
    'subparser_name': 'query',
    'koi': ['Brief Description', 'Last Business Case (BC) achieved'],
    "quarter": ["Q1 22/23", "Q4 21/22", "Q3 21/22"],
}
query_koi_milestones_cdg = {
    'test_name': 'query_koi_milestones',
    'subparser_name': 'query',
    'koi': "FBC CDG Approval",
    "quarter": ["Q1 22/23", "Q4 21/22", "Q3 21/22"],
}
query_koi_milestones_ipdc = {
    'test_name': 'query_koi_milestones',
    'subparser_name': 'query',
    'koi': ["OBC - IPDC Approval", "Planning Consents"],
    "quarter": ["Q4 21/22", "Q3 21/22", "Q2 21/22"],
}
q_koi_fn = {
    'test_name': 'query_koi_fn',
    'subparser_name': 'query',
    'koi_fn': 'test_query_keys',
    "quarter": ["Q1 22/23", "Q4 21/22", "Q3 21/22"],
}
q_koi_failure = {
    'test_name': 'query_koi_failure',
    'subparser_name': 'query',
}
if REPORTING_TYPE == 'cdg':
    QUERY_ARGS = [
        # q_koi_failure,
        q_koi_fn,
        query_koi_milestones_cdg,
        query_koi_quarters_cdg,
        q_koi_two_keys_cdg,
        q_koi_cdg
    ]
if REPORTING_TYPE == 'ipdc':
    QUERY_ARGS = [
        # q_koi_failure,
        q_koi_fn,
        query_koi_milestones_ipdc,
        query_koi_quarters_ipdc,
        q_koi_two_keys_ipdc,
        q_koi_ipdc,
    ]

if REPORTING_TYPE == 'ipdc':
    start_date = date_past(1)
    end_date = date_future(3)
if REPORTING_TYPE == 'cdg':
    start_date = date_past(6)
    end_date = date_future(6)

ms_standard = {
    "test_name": "ms_standard",
    "subparser_name": "milestones",
    "chart": chart,
    "dates": [start_date, end_date],
}
ms_groups_cdg = {
    "test_name": "ms_groups",
    "subparser_name": "milestones",
    "group": ["SCS"],
    "chart": chart,
}
ms_groups_ipdc = {
    "test_name": "ms_groups",
    "subparser_name": "milestones",
    "group": ["HSRG"],
    "chart": chart,
    "dates": [start_date, end_date],
}
ms_dates_ipdc = {
    "test_name": "ms_dates",
    "subparser_name": "milestones",
    "chart": chart,
    "dates": [start_date, end_date],
}
ms_blue_line_config = {
    "test_name": "ms_bl_config",
    "subparser_name": "milestones",
    "blue_line": "config_date",
    "chart": chart,
    "dates": [start_date, end_date],
}
ms_blue_line_today = {
    "test_name": "ms_bl_today",
    "subparser_name": "milestones",
    "blue_line": "today",
    "chart": chart,
    "dates": [start_date, end_date],
}
ms_koi_cdg = {
    "test_name": "ms_koi",
    "subparser_name": "milestones",
    "chart": chart,
    "koi": "FBC CDG Approval",
}
ms_koi_ipdc = {
    "test_name": "ms_koi",
    "subparser_name": "milestones",
    "chart": chart,
    "koi": "FBC - IPDC Approval",
}
ms_koi_fn = {
    "test_name": "ms_koi_fn",
    "subparser_name": "milestones",
    "chart": chart,
    "koi_fn": "milestone_keys",
    "dates": [start_date, end_date],
}
ms_quarters = {
    "test_name": "ms_quarters",
    "subparser_name": "milestones",
    "chart": chart,
    "quarter": ["Q4 21/22", "Q3 21/22", "Q2 21/22"],
    "dates": [start_date, end_date],
}
ms_stages = {
    "test_name": "ms_stages",
    "subparser_name": "milestones",
    "stage": [
        # 'Outline Business Case',
        'Full Business Case',
    ],
    "chart": chart,
    "dates": [start_date, end_date],

}
ms_koi_title_ipdc = {
    "test_name": "ms_koi",
    "subparser_name": "milestones",
    "chart": chart,
    "koi": "OBC - IPDC Approval",
    "title": "OBC Approvals",
}
# ms_for_reporting_cdg_far = {
#     "test_name": "ms_for_reporting_cdg_far",
#     "subparser_name": "milestones",
#     "blue_line": "config_date",
#     "chart": chart,
#     "dates": ["1/3/2023", "31/12/2024"],
#     "quarter": ["Q2 22/23", "Q1 22/23", "Q2 21/22"],
#
# }

if REPORTING_TYPE == 'ipdc':
    MILESTONES_OP_ARGS = [
        # ms_standard,  # to large / slow to test
        ms_koi_title_ipdc,
        ms_koi_fn,
        ms_koi_ipdc,
        ms_groups_ipdc,
        ms_dates_ipdc,
        ms_blue_line_config,
        ms_blue_line_today,
        ms_quarters,
        ms_stages,
    ]
if REPORTING_TYPE == 'cdg':
    MILESTONES_OP_ARGS = [
        # ms_for_reporting_cdg_far,  # only use for actual reporting
        ms_stages,
        ms_groups_cdg,
        ms_quarters,
        ms_koi_fn,
        ms_koi_cdg,
        ms_blue_line_today,
        ms_blue_line_config,
        ms_dates_ipdc,
        ms_groups_cdg,
        ms_standard,
    ]

dlion_standard = {
    "test_name": "dlion_standard",
    "subparser_name": "dandelion",
    "chart": chart,
}
dlion_pc_colour = {
    "test_name": "dlion_pc_colour",
    "subparser_name": "dandelion",
    "chart": chart,
    "pc": "A/R",
    # "angles": [250, 300, 360, 410, 440, 470, 500]
}
dlion_groups_cdg = {
    "test_name": "dlion_groups",
    "subparser_name": "dandelion",
    "group": ["SCS", "CFPD"],
    "chart": chart,
}
dlion_groups_ipdc = {
    "test_name": "dlion_groups",
    "subparser_name": "dandelion",
    "group": ["HSRG"],
    "chart": chart,
}
dlion_groups_ipdc_two = {
    "test_name": "dlion_groups_two",
    "subparser_name": "dandelion",
    "group": ["HSRG", "RIG"],
    "chart": chart,
}
dlion_stages_cdg = {
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
dlion_stages_ipdc = {
    "test_name": "dlion_stages",
    "subparser_name": "dandelion",
    "stage": [
        'Strategic Outline Case',
        'Outline Business Case',
        "Full Business Case",
    ],
    "chart": chart,
}
dlion_stages_abb_ipdc = {
    "test_name": "dlion_stages_abb_ipdc",
    "subparser_name": "dandelion",
    "stage": [
        'SOBC',
        'OBC',
        "FBC",
    ],
    "chart": chart,
}
dlion_stages_default_ipdc = {
    "test_name": "dlion_stages_default_ipdc",
    "subparser_name": "dandelion",
    "stage": [],
    "chart": chart,
}
dlion_quarter = {
    "test_name": "dlion_quarter",
    "subparser_name": "dandelion",
    "quarter": ["Q3 21/22"],
    "chart": chart,
}
dlion_angles_cdg = {
    "test_name": "dlion_angles",
    "subparser_name": "dandelion",
    "angles": [280, 360, 80],
    "chart": chart,
}
dlion_angles_ipdc = {
    "test_name": "dlion_angles",
    "subparser_name": "dandelion",
    "angles": [250, 305, 370, 60, 90, 130],
    "chart": chart,
}
# dlion_benefits = {
#     "test_name": "dlion_benefits",
#     "subparser_name": "dandelion",
#     "type": "benefits",
#     "chart": chart,
# }
dlion_remaining = {
    "test_name": "dlion_remaining",
    "subparser_name": "dandelion",
    "type": "remaining_costs",
    "chart": chart,
    # "angles": [250, 290, 350, 40, 90, 130, 170],
    # "order_by": "schedule",
    # "stage": [],
}
dlion_spent = {
    "test_name": "dlion_funded_resource",
    "subparser_name": "dandelion",
    "type": "spent",
    "chart": chart,
}
dlion_income = {
    "test_name": "dlion_income",
    "subparser_name": "dandelion",
    "type": "income",
    "chart": chart,
}
dlion_income_cdg = {
    "test_name": "dlion_income",
    "subparser_name": "dandelion",
    "type": "income",
    "chart": chart,
    "angles": [280, 360, 100],
}
dlion_funded_resource = {
    "test_name": "dlion_funded_resource",
    "subparser_name": "dandelion",
    "type": "funded_resource",
    "chart": chart,
}
dlion_ps_resource = {
    "test_name": "dlion_ps_resource",
    "subparser_name": "dandelion",
    "type": "ps_resource",
    "chart": chart,
}
dlion_contractor_resource = {
    "test_name": "dlion_contractor_resource",
    "subparser_name": "dandelion",
    "type": "contractor_resource",
    "chart": chart,
}
dlion_total_resource = {
    "test_name": "dlion_total_resource",
    "subparser_name": "dandelion",
    "type": "total_resource",
    "chart": chart,
}

# this will crash if project not in quarter data master
# dlion_cli_group_cdg = {
#     "test_name": "dlion_groups",
#     "subparser_name": "dandelion",
#     "group": ["WIT retail project", "Mayfield", "MSG"],
#     "chart": chart,
# }
# dlion_cli_group_ipdc = {
#     "test_name": "dlion_groups",
#     "subparser_name": "dandelion",
#     "group": ["RPE"],
#     "chart": chart,
# }
dlion_cli_pipeline_ipdc = {
    "test_name": "dlion_cli_pipeline_ipdc",
    "subparser_name": "dandelion",
    "group": ["pipeline"],
    "chart": chart,
}
dlion_pipeline_as_stage_ipdc = {
    "test_name": "dlion_pipeline_as_stage_ipdc",
    "subparser_name": "dandelion",
    "stage": ["pipeline", "SOBC"],
    "chart": chart,
}
dlion_stages_order_by_ipdc = {
    "test_name": "dlion_stages_order_by_ipdc",
    "subparser_name": "dandelion",
    "stage": [
        'SOBC',
        'OBC',
        "FBC",
    ],
    "order_by": "schedule",
    "chart": chart,
}
dlion_remove = {
    "test_name": "dlion_remove",
    "subparser_name": "dandelion",
    "chart": chart,
    "remove": ['NPR', 'A66'],
}
dlion_eviron_funds = {
    "test_name": "dlion_eviron_funds",
    "subparser_name": "dandelion",
    "chart": chart,
    "env_funds": [],
    # "angles": [300, 40],
}
dlion_eviron_funds_group = {
    "test_name": "dlion_eviron_funds",
    "subparser_name": "dandelion",
    "chart": chart,
    "env_funds": True,
    "angles": [320, 40],
}
dlion_near_spend = {
    "test_name": "dlion_near_spend",
    "subparser_name": "dandelion",
    "type": "near_spend",
    "chart": chart,
}

if REPORTING_TYPE == 'cdg':
    DANDELION_OP_ARGS_DICT = [
        # # dlion_cli_group,  # Failing.
        dlion_income_cdg,
        # # dlion_benefits,
        dlion_angles_cdg,
        dlion_quarter,
        dlion_stages_cdg,
        dlion_groups_cdg,
        dlion_standard,
    ]
if REPORTING_TYPE == 'ipdc':
    DANDELION_OP_ARGS_DICT = [
        dlion_pc_colour,
        dlion_near_spend,
        dlion_groups_ipdc_two,
        dlion_groups_ipdc,
        # dlion_eviron_funds,  ## failing due to text colour
        dlion_remove,
        dlion_total_resource,
        dlion_contractor_resource,
        dlion_ps_resource,
        dlion_funded_resource,
        dlion_spent,
        dlion_remaining,
        dlion_pipeline_as_stage_ipdc,
        dlion_stages_default_ipdc,
        dlion_stages_order_by_ipdc,
        dlion_cli_pipeline_ipdc,
        dlion_income,
        dlion_angles_ipdc,
        dlion_quarter,
        dlion_stages_ipdc,
        dlion_stages_abb_ipdc,
        dlion_standard,
    ]

# write some separate tests for dcas
sd_standard = {
    "test_name": "sd_standard",
    "subparser_name": "speed_dials",
}
sd_quarters = {
    "test_name": "sd_quarters",
    "subparser_name": "speed_dials",
    "quarter": ["Q2 21/22", "Q4 21/22"],
}
sd_too_many_quarters = {
    "test_name": "sd_quarters",
    "subparser_name": "speed_dials",
    "quarter": ["Q1 22/21", "Q4 21/22", "Q2 21/22"],
}
sd_stage = {
    "test_name": "sd_stage",
    "subparser_name": "speed_dials",
    "stage": [
        "OBC",
        "FBC",
    ],
}
sd_groups_cdg = {
    "test_name": "sd_groups",
    "subparser_name": "speed_dials",
    "group": ["SCS", "CFPD"],
}
sd_groups_ipdc = {
    "test_name": "sd_groups",
    "subparser_name": "speed_dials",
    "group": ["HSRG", "RLG"],
}

if REPORTING_TYPE == 'cdg':
    SPEED_DIAL_AND_DCA_OP_ARGS = [
        sd_groups_cdg,
        sd_stage,
        sd_quarters,
        sd_standard
    ]
if REPORTING_TYPE == 'ipdc':
    SPEED_DIAL_AND_DCA_OP_ARGS = [
        sd_too_many_quarters,
        sd_groups_ipdc,
        sd_stage,
        sd_quarters,
        sd_standard
    ]

pr_standard = {
    "test_name": "pr_standard",
    "subparser_name": "risks_portfolio",
}
pr_quarters = {
    "test_name": "pr_quarters",
    "subparser_name": "risks_portfolio",
    "quarter": ["Q1 22/23", "Q4 21/22"],
}
pr_groups_ipdc = {
    "test_name": "pr_groups",
    "subparser_name": "risks_portfolio",
    "group": ["HSRG", "AMS"],
}
pr_stage = {
    "test_name": "pr_stage",
    "subparser_name": "risks_portfolio",
    "stage": [
        "FBC",
    ],
}
pr_remove = {
    "test_name": "pr_remove",
    "subparser_name": "risks_portfolio",
    "group": ["HSRG"],
    "remove": ["HS2 Ph 1"]
}

PORT_RISK_OP_ARGS =[
    pr_remove,
    pr_stage,
    pr_groups_ipdc,
    pr_quarters,
    pr_standard,
]

PORT_RISK_OP_WORD_ARGS = [
    pr_standard,
    pr_quarters,
]

sum_standard = {
    "test_name": "sum_standard",
    "subparser_name": "summaries",
    "type": "short",
}

sum_group = {
    "test_name": "sum_standard",
    "subparser_name": "summaries",
    "type": "short",
    "group": ["HSRG"],
}

sum_group_two = {
    "test_name": "sum_standard",
    "subparser_name": "summaries",
    "type": "short",
    "group": ["A358"],
}

SUM_OP_ARGS = [
    # sum_standard,
    sum_group,
    # sum_group_two,
]


# 'rag_number': '5',
# 'order_by': 'cost',
# 'none_handle': 'none',
