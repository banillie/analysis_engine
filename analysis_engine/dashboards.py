from dateutil import parser
from openpyxl import Workbook
from openpyxl.formatting.rule import IconSetRule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting import Rule

from analysis_engine.dictionaries import (
    DATA_KEY_DICT,
    BC_STAGE_DICT_FULL_TO_ABB,
    CONVERT_RAG,
    rag_txt_list,
    conf_list,
    risk_list,
    DASHBOARD_KEYS,
    DCA_KEYS,
    STANDARDISE_COST_KEYS,
)
from analysis_engine.dandelion import dandelion_number_text
from analysis_engine.colouring import black_text, fill_colour_list
from analysis_engine.milestones import MilestoneData, get_milestone_date
from analysis_engine.segmentation import get_group
from analysis_engine.error_msgs import InputError
from analysis_engine.costs import CostData, convert_none_types
from analysis_engine.render_utils import plus_minus_days
from analysis_engine.settings import get_remove_income


def narrative_dashboard(master, wb: Workbook) -> None:
    ws = wb.active

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name in master["current_projects"]:
            # Group
            ws.cell(row=row_num, column=2).value = master["project_information"][
                project_name
            ]["Directorate"]
            # Abbreviation
            ws.cell(row=row_num, column=4).value = master["project_information"][
                project_name
            ]["Abbreviations"]
            # Stage
            bc_stage = master["master_data"][0]["data"][project_name][
                DATA_KEY_DICT["IPDC approval point"]
            ]
            ws.cell(row=row_num, column=5).value = BC_STAGE_DICT_FULL_TO_ABB[bc_stage]
            costs = master["master_data"][0]["data"][project_name][
                DATA_KEY_DICT["Total Forecast"]
            ]
            ws.cell(row=row_num, column=6).value = dandelion_number_text(costs)

            overall_dca = CONVERT_RAG[
                master["master_data"][0]["data"][project_name][
                    DATA_KEY_DICT["Departmental DCA"]
                ]
            ]
            ws.cell(row=row_num, column=7).value = overall_dca
            if overall_dca == "None":
                ws.cell(row=row_num, column=7).value = ""

            sro_n = master["master_data"][0]["data"][project_name]["SRO Narrative"]
            ws.cell(row=row_num, column=8).value = sro_n

        """list of columns with conditional formatting"""
        list_columns = ["g"]

        """same loop but the text is black. In addition these two loops go 
        through the list_columns list above"""
        for column in list_columns:
            for i, dca in enumerate(rag_txt_list):
                text = black_text
                fill = fill_colour_list[i]
                dxf = DifferentialStyle(font=text, fill=fill)
                rule = Rule(
                    type="containsText", operator="containsText", text=dca, dxf=dxf
                )
                for_rule_formula = 'NOT(ISERROR(SEARCH("' + dca + '",' + column + "5)))"
                rule.formula = [for_rule_formula]
                ws.conditional_formatting.add(column + "5:" + column + "60", rule)

        for row_num in range(2, ws.max_row + 1):
            for col_num in range(5, ws.max_column + 1):
                if ws.cell(row=row_num, column=col_num).value == 0:
                    ws.cell(row=row_num, column=col_num).value = "-"

    # return wb


def cdg_dashboard(master, wb: Workbook) -> None:
    ws = wb.active

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name in master["current_projects"]:
            ws.cell(row=row_num, column=2).value = master["project_information"][
                project_name
            ]["Directorate"]
            ws.cell(row=row_num, column=4).value = master["project_information"][
                project_name
            ]["Abbreviations"]
            bc_stage = master["master_data"][0]["data"][project_name][
                DATA_KEY_DICT["IPDC approval point"]
            ]
            ws.cell(row=row_num, column=5).value = BC_STAGE_DICT_FULL_TO_ABB[bc_stage]
            costs = master["master_data"][0]["data"][project_name][
                DATA_KEY_DICT["Total Forecast"]
            ]
            ws.cell(row=row_num, column=6).value = dandelion_number_text(
                costs, none_handle="none"
            )
            income = master["master_data"][0]["data"][project_name]["Total Income"]
            ws.cell(row=row_num, column=7).value = dandelion_number_text(
                income, none_handle="none"
            )
            benefits = master["master_data"][0]["data"][project_name]["Total Benefits"]
            ws.cell(row=row_num, column=8).value = dandelion_number_text(
                benefits, none_handle="none"
            )
            vfm = master["master_data"][0]["data"][project_name]["VfM Category"]
            ws.cell(row=row_num, column=9).value = vfm
            overall_dca = CONVERT_RAG[
                master["master_data"][0]["data"][project_name][
                    DATA_KEY_DICT["Departmental DCA"]
                ]
            ]
            ws.cell(row=row_num, column=10).value = overall_dca
            if overall_dca == "None":
                ws.cell(row=row_num, column=10).value = ""

            for i, key in enumerate(conf_list):
                dca = CONVERT_RAG[master["master_data"][0]["data"][project_name][key]]
                ws.cell(row=row_num, column=11 + i).value = dca

            for i, key in enumerate(risk_list):
                risk = master["master_data"][0]["data"][project_name][key]
                if risk == "YES":
                    ws.cell(row=row_num, column=14 + i).value = risk

        """list of columns with conditional formatting"""
        list_columns = ["j", "k", "l", "m"]

        """same loop but the text is black. In addition these two loops 
        go through the list_columns list above"""
        for column in list_columns:
            for i, dca in enumerate(rag_txt_list):
                text = black_text
                fill = fill_colour_list[i]
                dxf = DifferentialStyle(font=text, fill=fill)
                rule = Rule(
                    type="containsText", operator="containsText", text=dca, dxf=dxf
                )
                for_rule_formula = 'NOT(ISERROR(SEARCH("' + dca + '",' + column + "5)))"
                rule.formula = [for_rule_formula]
                ws.conditional_formatting.add(column + "5:" + column + "60", rule)

        for row_num in range(2, ws.max_row + 1):
            for col_num in range(5, ws.max_column + 1):
                if ws.cell(row=row_num, column=col_num).value == 0:
                    ws.cell(row=row_num, column=col_num).value = "-"

    # return wb


def ipdc_dashboard(md, wb: Workbook, **op_args) -> Workbook:
    financial_dashboard(md, wb, **op_args)
    # resource_dashboard(md, wb, **op_args)
    #
    # milestones = MilestoneData(md, **op_args)
    # m_filtered = MilestoneData(md, **op_args)
    # m_filtered.filter_chart_info(type=["Approval"])
    # schedule_dashboard(md, milestones, m_filtered, wb)
    # benefits_dashboard(md, wb)
    #
    # overall_dashboard(md, milestones, wb, **op_args)

    return wb


def resource_dashboard(master, wb: Workbook, **kwargs) -> Workbook:
    ws = wb["Resource"]

    current_data = master.master_data[0]["data"]
    last_data = master.master_data[1]["data"]
    last_qrt_group = get_group(kwargs["group"], master)

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name in master.current_projects:
            if project_name in last_qrt_group:
                kwargs["group"] = [project_name]
                kwargs["quarter"] = ["standard"]
            else:
                kwargs["group"] = [project_name]
                kwargs["quarter"] = [str(master.current_quarter)]

            """BC Stage"""
            bc_stage = current_data[project_name]["IPDC approval point"]
            ws.cell(row=row_num, column=4).value = BC_STAGE_DICT_FULL_TO_ABB(bc_stage)

            "Resourcing data"
            resource_keys = [
                "DfTc Public Sector Employees",
                "DfTc External Contractors",
                "DfTc Project Team Total",
                "DfTc Funded Posts",
                "DfTc Resource Gap",
                "DfTc Resource Gap Criticality",
            ]

            for i, key in enumerate(resource_keys):
                try:
                    if key == "DfTc Resource Gap Criticality":
                        ws.cell(row=row_num, column=5 + i).value = CONVERT_RAG(
                            current_data[project_name][key]
                        )
                    else:
                        ws.cell(row=row_num, column=5 + i).value = current_data[
                            project_name
                        ][key]
                except KeyError:
                    raise InputError(
                        key + " key is not in quarter master. This key must"
                        " be present for dashboard compilation. Stopping. "
                        "Make sure all resource keys are in Master."
                    )

            """DCA rating - this quarter"""
            ws.cell(row=row_num, column=12).value = CONVERT_RAG(
                current_data[project_name]["Overall Resource DCA - Now"]
            )
            """DCA rating - last qrt"""
            try:
                ws.cell(row=row_num, column=13).value = CONVERT_RAG(
                    last_data[project_name]["Overall Resource DCA - Now"]
                )
            except KeyError:
                ws.cell(row=row_num, column=13).value = ""
            """DCA rating - 2 qrts ago"""
            try:
                ws.cell(row=row_num, column=14).value = CONVERT_RAG(
                    master.master_data[2]["data"][project_name][
                        "Overall Resource DCA - Now"
                    ]
                )
            except (KeyError, IndexError):
                ws.cell(row=row_num, column=14).value = ""
            """DCA rating - 3 qrts ago"""
            try:
                ws.cell(row=row_num, column=15).value = CONVERT_RAG(
                    master.master_data[3]["data"][project_name][
                        "Overall Resource DCA - Now"
                    ]
                )
            except (KeyError, IndexError):
                ws.cell(row=row_num, column=15).value = ""

    """list of columns with conditional formatting"""
    list_columns = ["j", "l", "m", "n", "o"]

    """same loop but the text is black. In addition these two loops go through the list_columns list above"""
    for column in list_columns:
        for i, dca in enumerate(rag_txt_list):
            text = black_text
            fill = fill_colour_list[i]
            dxf = DifferentialStyle(font=text, fill=fill)
            rule = Rule(type="containsText", operator="containsText", text=dca, dxf=dxf)
            for_rule_formula = 'NOT(ISERROR(SEARCH("' + dca + '",' + column + "5)))"
            rule.formula = [for_rule_formula]
            ws.conditional_formatting.add("" + column + "5:" + column + "60", rule)

    # for row_num in range(2, ws.max_row + 1):
    #     for col_num in range(5, ws.max_column+1):
    #         if ws.cell(row=row_num, column=col_num).value == 0:
    #             ws.cell(row=row_num, column=col_num).value = '-'

    return wb


def financial_dashboard(
    md,
    wb: Workbook,
    **op_args,
) -> Workbook:

    ws = wb["Finance"]
    cmd = md["master_data"][0]["data"]  # cmd = current master data
    lmd = md["master_data"][1]["data"]

    rm = get_remove_income(op_args)

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name not in md["current_projects"]:
            continue

        """BC Stage"""
        bc_stage = cmd[project_name][DASHBOARD_KEYS["BC_STAGE"]]
        ws.cell(row=row_num, column=4).value = BC_STAGE_DICT_FULL_TO_ABB[bc_stage]
        """Total WLC"""
        wlc_now = convert_none_types(
            cmd[project_name][STANDARDISE_COST_KEYS[op_args["report"]]["total"]]
        )
        if project_name in rm:
            wlc_now = wlc_now - convert_none_types(
                cmd[project_name][
                    STANDARDISE_COST_KEYS[op_args["report"]]["income_total"]
                ]
            )
        ws.cell(row=row_num, column=6).value = wlc_now
        """WLC variance against lst quarter"""
        try:
            wlc_lst_qrt = convert_none_types(
                lmd[project_name][STANDARDISE_COST_KEYS[op_args["report"]]["total"]]
            )
            if project_name in rm:
                wlc_lst_qrt = wlc_lst_qrt - convert_none_types(
                    lmd[project_name][
                        STANDARDISE_COST_KEYS[op_args["report"]]["income_total"]
                    ]
                )
            print(project_name, wlc_now, wlc_lst_qrt)
            diff_lst_qrt = wlc_now - wlc_lst_qrt
            if float(diff_lst_qrt) > 0.49 or float(diff_lst_qrt) < -0.49:
                ws.cell(row=row_num, column=7).value = diff_lst_qrt
            else:
                ws.cell(row=row_num, column=7).value = "-"
        except KeyError:
            ws.cell(row=row_num, column=7).value = "-"

        # """WLC variance against baseline quarter"""
        # wlc_baseline = costs.baseline[str(md.current_quarter)]['total']
        # try:
        #     diff_bl = wlc_now - wlc_baseline
        #     if float(diff_bl) > 0.49 or float(diff_bl) < -0.49:
        #         ws.cell(row=row_num, column=8).value = diff_bl
        #     else:
        #         ws.cell(row=row_num, column=8).value = "-"
        # # exception is here as some projects e.g. Hs2 phase 2b have (real) written into historical totals
        # except TypeError:
        #     pass

        con = cmd[project_name][DASHBOARD_KEYS["CONTINGENCY"]]
        if con == 0 or con is None:
            con = "-"
        ws.cell(row=row_num, column=13).value = con

        """OB"""
        ob = cmd[project_name][DASHBOARD_KEYS["OB"]]
        if ob == 0 or ob is None:
            ob = "-"
        ws.cell(row=row_num, column=14).value = ob

        """financial DCA ratings"""
        for i, q in enumerate(md["quarter_list"]):
            try:
                ws.cell(row=row_num, column=15 + i).value = CONVERT_RAG[
                    md["master_data"][i]["data"][project_name][
                        DCA_KEYS[op_args["report"]]["finance"]
                    ]
                ]
            except KeyError:
                ws.cell(row=row_num, column=16).value = ""

    """list of columns with conditional formatting"""
    list_columns = ["o", "p", "q", "r", "s"]

    """same loop but the text is black. In addition these two loops go through the list_columns list above"""
    for column in list_columns:
        for i, dca in enumerate(rag_txt_list):
            text = black_text
            fill = fill_colour_list[i]
            dxf = DifferentialStyle(font=text, fill=fill)
            rule = Rule(type="containsText", operator="containsText", text=dca, dxf=dxf)
            for_rule_formula = 'NOT(ISERROR(SEARCH("' + dca + '",' + column + "5)))"
            rule.formula = [for_rule_formula]
            ws.conditional_formatting.add("" + column + "5:" + column + "60", rule)

    return wb


def schedule_dashboard(
    master,
    milestones: MilestoneData,
    m_filtered,
    wb: Workbook,
    **op_args,
) -> Workbook:
    ws = wb["Schedule"]
    # overall_ws = wb.worksheets[3]

    current_data = master.master_data[0]["data"]
    last_data = master.master_data[1]["data"]
    IPDC_DATE = parser.parse(op_args["data"], dayfirst=True).date()
    #     get_ipdc_date(str(root_path) + "/core_data/ipdc_config.ini", "ipdc_date"),
    #     dayfirst=True,
    # ).date()

    def get_next_milestone(p_name: str, mils: MilestoneData) -> list:
        for x in mils.milestone_dict[milestones.iter_list[0]].values():
            if x["Project"] == p_name:
                d = x["Date"]
                ms = x["Milestone"]
                if d > IPDC_DATE:
                    return [ms, d]

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name in master.current_projects:
            """IPDC approval point"""
            bc_stage = current_data[project_name]["IPDC approval point"]
            ws.cell(row=row_num, column=4).value = BC_STAGE_DICT_FULL_TO_ABB(bc_stage)

            """Next milestone name and variance"""

            abb = master.abbreviations[project_name]["abb"]
            try:
                g = get_next_milestone(abb, m_filtered)
                milestone = g[0]
                date = g[1]
                ws.cell(row=row_num, column=6).value = milestone
                ws.cell(row=row_num, column=7).value = date

                # lq_date = get_milestone_date(
                #     abb, milestones.milestone_dict, "last", " " + milestone
                # )
                # try:
                #     change = (date - lq_date).days
                #     ws.cell(row=row_num, column=8).value = plus_minus_days(change)
                #     if change > 25:
                #         ws.cell(row=row_num, column=8).font = Font(
                #             name="Arial", size=10, color="00fc2525"
                #         )
                # except TypeError:
                #     pass
                #     # ws.cell(row=row_num, column=8).value = ""

                # bl_date = get_milestone_date(
                #     abb, milestones.milestone_dict, "bl_one", " " + milestone
                # )
                # try:
                #     change = (date - bl_date).days
                #     ws.cell(row=row_num, column=9).value = plus_minus_days(change)
                #     if change > 25:
                #         ws.cell(row=row_num, column=9).font = Font(
                #             name="Arial", size=10, color="00fc2525"
                #         )
                # except TypeError:
                #     pass
            except TypeError:
                ws.cell(row=row_num, column=6).value = "None"
                ws.cell(row=row_num, column=7).value = None

            milestone_keys = [
                "Start of Construction/build",
                "Start of Operation",
                "Full Operations",
                "Project End Date",
            ]  # code legency needs a space at start of keys
            add_column = 0
            for m in milestone_keys:
                abb = master.abbreviations[project_name]["abb"]
                current = get_milestone_date(
                    abb, milestones.milestone_dict, str(master.current_quarter), m
                )
                last_quarter = get_milestone_date(
                    abb, milestones.milestone_dict, str(master.quarter_list[1]), m
                )
                # bl = get_milestone_date(abb, milestones.milestone_dict, "bl_one", m)
                # if current == None:
                #     current = "None"
                ws.cell(row=row_num, column=10 + add_column).value = current
                if current is not None and current < IPDC_DATE:
                    # if m == "Full Operations":
                    #     overall_ws.cell(row=row_num, column=9).value = "Completed"
                    ws.cell(row=row_num, column=10 + add_column).value = "Completed"
                if current is None:
                    ws.cell(row=row_num, column=10 + add_column).value = "None"
                try:
                    last_change = (current - last_quarter).days
                    # if m == "Full Operations":
                    #     ws.cell(
                    #         row=row_num, column=10).value = plus_minus_days(last_change)
                    ws.cell(
                        row=row_num, column=11 + add_column
                    ).value = plus_minus_days(last_change)
                    # if last_change is not None and last_change > 46:
                    #     # if m == "Full Operations":
                    #     #     overall_ws.cell(row=row_num, column=10).font = Font(
                    #     #         name="Arial", size=10, color="00fc2525"
                    #     #     )
                    #     ws.cell(row=row_num, column=11 + add_column).font = Font(
                    #         name="Arial", size=10, color="00fc2525"
                    #     )
                except TypeError:
                    pass
                # try:
                #     bl_change = (current - bl).days
                #     # if m == "Full Operations":
                #     #     overall_ws.cell(
                #     #         row=row_num, column=11
                #     #     ).value = plus_minus_days(bl_change)
                #     ws.cell(
                #         row=row_num, column=12 + add_column
                #     ).value = plus_minus_days(bl_change)
                #     if bl_change is not None and bl_change > 85:
                #         # if m == "Full Operations":
                #         #     overall_ws.cell(row=row_num, column=11).font = Font(
                #         #         name="Arial", size=10, color="00fc2525"
                #         #     )
                #         ws.cell(row=row_num, column=12 + add_column).font = Font(
                #             name="Arial", size=10, color="00fc2525"
                #         )
                # except TypeError:
                #     pass
                add_column += 3

            """schedule DCA rating - this quarter"""
            ws.cell(row=row_num, column=22).value = CONVERT_RAG(
                current_data[project_name]["SRO Schedule Confidence"]
            )
            """schedule DCA rating - last qrt"""
            try:
                ws.cell(row=row_num, column=23).value = CONVERT_RAG(
                    last_data[project_name]["SRO Schedule Confidence"]
                )
            except KeyError:
                ws.cell(row=row_num, column=23).value = ""
            """schedule DCA rating - 2 qrts ago"""
            try:
                ws.cell(row=row_num, column=24).value = CONVERT_RAG(
                    master.master_data[2]["data"][project_name][
                        "SRO Schedule Confidence"
                    ]
                )
            except (KeyError, IndexError):
                ws.cell(row=row_num, column=24).value = ""
            """schedule DCA rating - 3 qrts ago"""
            try:
                ws.cell(row=row_num, column=25).value = CONVERT_RAG(
                    master.master_data[3]["data"][project_name][
                        "SRO Schedule Confidence"
                    ]
                )
            except (KeyError, IndexError):
                ws.cell(row=row_num, column=25).value = ""
            # """schedule DCA rating - baseline"""
            # bl_i = master.bl_index["ipdc_milestones"][project_name][2]
            # try:
            #     ws.cell(row=row_num, column=26).value = CONVERT_RAG(
            #         master.master_data[bl_i]["data"][project_name][
            #             "SRO Schedule Confidence"
            #         ]
            #     )
            # except KeyError:  # schedule confidence key not in all masters.
            #     pass

    """list of columns with conditional formatting"""
    list_columns = ["v", "w", "x", "y", "z"]

    """same loop but the text is black. In addition these two loops go through the list_columns list above"""
    for column in list_columns:
        for i, dca in enumerate(rag_txt_list):
            text = black_text
            fill = fill_colour_list[i]
            dxf = DifferentialStyle(font=text, fill=fill)
            rule = Rule(type="containsText", operator="containsText", text=dca, dxf=dxf)
            for_rule_formula = 'NOT(ISERROR(SEARCH("' + dca + '",' + column + "5)))"
            rule.formula = [for_rule_formula]
            ws.conditional_formatting.add("" + column + "5:" + column + "60", rule)

    for row_num in range(2, ws.max_row + 1):
        for col_num in range(5, ws.max_column + 1):
            if ws.cell(row=row_num, column=col_num).value == 0:
                ws.cell(row=row_num, column=col_num).value = "-"

    return wb


def benefits_dashboard(master, wb: Workbook) -> Workbook:
    ws = wb["Benefits_VfM"]
    # overall_ws = wb.worksheets[3]

    current_data = master.master_data[0]["data"]
    last_data = master.master_data[1]["data"]

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name in master.current_projects:
            # bl_i = master.bl_index["ipdc_benefits"][project_name][2]
            # baseline_data = master.master_data[bl_i]["data"]

            """BICC approval point"""
            bc_stage = current_data[project_name]["IPDC approval point"]
            ws.cell(row=row_num, column=4).value = BC_STAGE_DICT_FULL_TO_ABB(bc_stage)
            # try:
            #     bc_stage_lst_qrt = last_data[project_name]["IPDC approval point"]
            #     if bc_stage != bc_stage_lst_qrt:
            #         ws.cell(row=row_num, column=4).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # except KeyError:
            #     pass
            # """Next stage"""
            # proj_stage = current_data[project_name]["Project stage"]
            # ws.cell(row=row_num, column=5).value = proj_stage
            # try:
            #     proj_stage_lst_qrt = last_data[project_name]["Project stage"]
            #     if proj_stage != proj_stage_lst_qrt:
            #         ws.cell(row=row_num, column=5).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # except KeyError:
            #     pass

            """initial bcr"""
            initial_bcr = current_data[project_name][
                "Initial Benefits Cost Ratio (BCR)"
            ]
            ws.cell(row=row_num, column=6).value = initial_bcr
            # """initial bcr baseline"""
            # try:
            # baseline_initial_bcr = baseline_data[project_name][
            #     "Initial Benefits Cost Ratio (BCR)"
            # ]
            # if baseline_initial_bcr != 0:
            #     ws.cell(row=row_num, column=7).value = baseline_initial_bcr
            # else:
            #     ws.cell(row=row_num, column=7).value = ""
            # if initial_bcr != baseline_initial_bcr:
            #     if baseline_initial_bcr is None:
            #         pass
            #     else:
            #         ws.cell(row=row_num, column=6).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            #         ws.cell(row=row_num, column=7).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # # except TypeError:
            # #     ws.cell(row=row_num, column=7).value = ""

            """adjusted bcr"""
            adjusted_bcr = current_data[project_name][
                "Adjusted Benefits Cost Ratio (BCR)"
            ]
            ws.cell(row=row_num, column=8).value = adjusted_bcr
            # """adjusted bcr baseline"""
            # # try:
            # baseline_adjusted_bcr = baseline_data[project_name][
            #     "Adjusted Benefits Cost Ratio (BCR)"
            # ]
            # if baseline_adjusted_bcr != 0:
            #     ws.cell(row=row_num, column=9).value = baseline_adjusted_bcr
            # else:
            #     ws.cell(row=row_num, column=9).value = ""
            # if adjusted_bcr != baseline_adjusted_bcr:
            #     if baseline_adjusted_bcr is not None:
            #         ws.cell(row=row_num, column=8).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            #         ws.cell(row=row_num, column=9).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # # except TypeError:
            # #     ws.cell(row=row_num, column=9).value = ""

            """vfm category now"""
            if current_data[project_name]["VfM Category single entry"] is None:
                vfm_cat = (
                    str(current_data[project_name]["VfM Category lower range"])
                    + " - "
                    + str(current_data[project_name]["VfM Category upper range"])
                )
                if vfm_cat == "None - None":
                    vfm_cat = "None"
                ws.cell(row=row_num, column=10).value = vfm_cat
                # overall_ws.cell(row=row_num, column=8).value = vfm_cat

            else:
                vfm_cat = current_data[project_name]["VfM Category single entry"]
                ws.cell(row=row_num, column=10).value = vfm_cat
                # overall_ws.cell(row=row_num, column=8).value = vfm_cat

            # """vfm category baseline"""
            # try:
            #     if baseline_data[project_name]["VfM Category single entry"] is None:
            #         vfm_cat_baseline = (
            #                 str(baseline_data[project_name]["VfM Category lower range"])
            #                 + " - "
            #                 + str(baseline_data[project_name]["VfM Category upper range"])
            #         )
            #         ws.cell(row=row_num, column=11).value = vfm_cat_baseline
            #     else:
            #         vfm_cat_baseline = baseline_data[project_name][
            #             "VfM Category single entry"
            #         ]
            #         ws.cell(row=row_num, column=11).value = vfm_cat_baseline
            #
            # except KeyError:
            #     try:
            #         vfm_cat_baseline = baseline_data[project_name][
            #             "VfM Category single entry"
            #         ]
            #         ws.cell(row=row_num, column=11).value = vfm_cat_baseline
            #     except KeyError:
            #         vfm_cat_baseline = baseline_data[project_name]["VfM Category"]
            #         ws.cell(row=row_num, column=11).value = vfm_cat_baseline
            #
            # if vfm_cat != vfm_cat_baseline:
            #     if vfm_cat_baseline is None:
            #         pass
            #     else:
            #         ws.cell(row=row_num, column=10).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            #         ws.cell(row=row_num, column=11).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            #         # overall_ws.cell(row=row_num, column=8).font = Font(
            #         #     name="Arial", size=10, color="00fc2525"
            #         # )

            """total monetised benefits"""
            tmb = current_data[project_name][
                "Total BEN Forecast - Total Monetised Benefits"
            ]
            ws.cell(row=row_num, column=12).value = tmb
            # """tmb variance"""
            # baseline_tmb = baseline_data[project_name][
            #     "Total BEN Forecast - Total Monetised Benefits"
            # ]
            # tmb_variance = tmb - baseline_tmb
            # ws.cell(row=row_num, column=13).value = tmb_variance
            # if tmb_variance == 0:
            #     ws.cell(row=row_num, column=13).value = "-"
            # try:
            #     percentage_change = ((tmb - baseline_tmb) / tmb) * 100
            #     if percentage_change > 5 or percentage_change < -5:
            #         ws.cell(row=row_num, column=13).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # except ZeroDivisionError:
            #     pass

            # In year benefits
            iyb = current_data[project_name]["BEN Forecast In-Year"]
            ws.cell(row=row_num, column=14).value = iyb
            # try:
            #     iyb_bl = baseline_data[project_name]["BEN Forecast In-Year"]
            #     iyb_diff = iyb - iyb_bl
            #     ws.cell(row=row_num, column=15).value = iyb_diff
            #     if iyb_diff == 0:
            #         ws.cell(row=row_num, column=15).value = "-"
            #     percentage_change = ((iyb - iyb_bl) / iyb) * 100
            #     if percentage_change > 5 or percentage_change < -5:
            #         ws.cell(row=row_num, column=15).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # except (KeyError, ZeroDivisionError):  # key only present from Q2 20/21
            #     pass

            """benefits DCA rating - this quarter"""
            ws.cell(row=row_num, column=16).value = CONVERT_RAG(
                current_data[project_name]["SRO Benefits RAG"]
            )
            """benefits DCA rating - last qrt"""
            try:
                ws.cell(row=row_num, column=17).value = CONVERT_RAG(
                    last_data[project_name]["SRO Benefits RAG"]
                )
            except KeyError:
                ws.cell(row=row_num, column=17).value = ""
            """benefits DCA rating - 2 qrts ago"""
            try:
                ws.cell(row=row_num, column=18).value = CONVERT_RAG(
                    master.master_data[2]["data"][project_name]["SRO Benefits RAG"]
                )
            except (KeyError, IndexError):
                ws.cell(row=row_num, column=18).value = ""
            """benefits DCA rating - 3 qrts ago"""
            try:
                ws.cell(row=row_num, column=19).value = CONVERT_RAG(
                    master.master_data[3]["data"][project_name]["SRO Benefits RAG"]
                )
            except (KeyError, IndexError):
                ws.cell(row=row_num, column=19).value = ""

            # """benefits DCA rating - baseline"""
            # ws.cell(row=row_num, column=20).value = CONVERT_RAG(
            #     baseline_data[project_name]["SRO Benefits RAG"]
            # )

    """list of columns with conditional formatting"""
    list_columns = ["p", "q", "r", "s", "t"]

    """loops below place conditional formatting (cf) rules into the wb. There are two as the dashboard currently has
    two distinct sections/headings, which do not require cf. Therefore, cf starts and ends at the stated rows. this
    is hard code that will need to be changed should the position of information in the dashboard change. It is an
    easy change however"""

    """same loop but the text is black. In addition these two loops go through the list_columns list above"""
    for column in list_columns:
        for i, dca in enumerate(rag_txt_list):
            text = black_text
            fill = fill_colour_list[i]
            dxf = DifferentialStyle(font=text, fill=fill)
            rule = Rule(type="containsText", operator="containsText", text=dca, dxf=dxf)
            for_rule_formula = 'NOT(ISERROR(SEARCH("' + dca + '",' + column + "5)))"
            rule.formula = [for_rule_formula]
            ws.conditional_formatting.add("" + column + "5:" + column + "60", rule)

    # for row_num in range(2, ws.max_row + 1):
    #     for col_num in range(5, ws.max_column+1):
    #         if ws.cell(row=row_num, column=col_num).value == 0:
    #             ws.cell(row=row_num, column=col_num).value = '-'

    return wb


def overall_dashboard(
    master, milestones: MilestoneData, wb: Workbook, **op_args
) -> Workbook:
    ws = wb["Overall"]

    current_data = master.master_data[0]["data"]
    last_data = master.master_data[1]["data"]
    last_qrt_group = get_group(op_args["group"], master, 1)

    IPDC_DATE = parser.parse(op_args["data"], dayfirst=True).date()
    #     get_ipdc_date(str(root_path) + "/core_data/ipdc_config.ini", "ipdc_date"),
    #     dayfirst=True,
    # ).date()

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=2).value
        if project_name in master.current_projects:
            if project_name in last_qrt_group:
                op_args["group"] = [project_name]
                op_args["quarter"] = ["standard"]
            else:
                op_args["group"] = [project_name]
                op_args["quarter"] = [str(master.current_quarter)]
            c = CostData(master, **op_args)
            """BC Stage"""
            bc_stage = current_data[project_name]["IPDC approval point"]
            # ws.cell(row=row_num, column=4).value = BC_STAGE_DICT_FULL_TO_ABB(bc_stage)
            ws.cell(row=row_num, column=3).value = BC_STAGE_DICT_FULL_TO_ABB(bc_stage)
            # try:
            #     bc_stage_lst_qrt = last_data[project_name]["IPDC approval point"]
            #     if bc_stage != bc_stage_lst_qrt:
            #         # ws.cell(row=row_num, column=4).font = Font(
            #         #     name="Arial", size=10, color="00fc2525"
            #         # )
            #         ws.cell(row=row_num, column=3).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # except KeyError:
            #     pass

            # """planning stage"""
            # plan_stage = current_data[project_name]["Project stage"]
            # # ws.cell(row=row_num, column=5).value = plan_stage
            # ws.cell(row=row_num, column=4).value = plan_stage
            # try:
            #     plan_stage_lst_qrt = last_data[project_name]["Project stage"]
            #     if plan_stage != plan_stage_lst_qrt:
            #         # ws.cell(row=row_num, column=5).font = Font(
            #         #     name="Arial", size=10, color="00fc2525"
            #         # )
            #         ws.cell(row=row_num, column=4).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # except KeyError:
            #     pass

            """Total WLC"""
            wlc_now = c.c_totals[str(master.current_quarter)]["total"]
            # ws.cell(row=row_num, column=6).value = wlc_now
            ws.cell(row=row_num, column=5).value = wlc_now
            """WLC variance against lst quarter"""
            try:
                lst_qrt_costs = c.c_totals[str(master.quarter_list[1])]["total"]
                diff_lst_qrt = wlc_now - lst_qrt_costs
                if float(diff_lst_qrt) > 0.49 or float(diff_lst_qrt) < -0.49:
                    # ws.cell(row=row_num, column=7).value = diff_lst_qrt
                    ws.cell(row=row_num, column=6).value = diff_lst_qrt
                else:
                    # ws.cell(row=row_num, column=7).value = "-"
                    ws.cell(row=row_num, column=6).value = "-"

                # try:
                #     percentage_change = ((wlc_now - lst_qrt_costs) / wlc_now) * 100
                #     if percentage_change > 5 or percentage_change < -5:
                #         # ws.cell(row=row_num, column=7).font = Font(
                #         #     name="Arial", size=10, color="00fc2525"
                #         # )
                #         ws.cell(row=row_num, column=6).font = Font(
                #             name="Arial", size=10, color="00fc2525"
                #         )
                # except ZeroDivisionError:
                #     pass

            except KeyError:
                ws.cell(row=row_num, column=6).value = "-"

            """WLC variance against baseline"""
            wlc_baseline = c.c_bl_totals[str(master.current_quarter)]["total"]
            # bl = master.bl_index["ipdc_costs"][project_name][2]
            # wlc_baseline = master.master_data[bl]["data"][project_name][
            #     "Total Forecast"
            # ]
            try:
                diff_bl = wlc_now - wlc_baseline
                if float(diff_bl) > 0.49 or float(diff_bl) < -0.49:
                    # ws.cell(row=row_num, column=8).value = diff_bl
                    ws.cell(row=row_num, column=7).value = diff_bl
                else:
                    # ws.cell(row=row_num, column=8).value = "-"
                    ws.cell(row=row_num, column=7).value = "-"
            except TypeError:  # exception is here as some projects e.g. Hs2 phase 2b have (real) written into historical totals
                pass

            # try:
            #     percentage_change = ((wlc_now - wlc_baseline) / wlc_now) * 100
            #     if percentage_change > 5 or percentage_change < -5:
            #         # ws.cell(row=row_num, column=8).font = Font(
            #         #     name="Arial", size=10, color="00fc2525"
            #         # )
            #         ws.cell(row=row_num, column=7).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            #
            # except (
            #         ZeroDivisionError,
            #         TypeError,
            # ):  # zerodivision error obvious, type error handling as above
            #     pass

            """vfm category now"""
            if current_data[project_name]["VfM Category single entry"] is None:
                vfm_cat = (
                    str(current_data[project_name]["VfM Category lower range"])
                    + " - "
                    + str(current_data[project_name]["VfM Category upper range"])
                )
                if vfm_cat == "None - None":
                    vfm_cat = "None"
                # ws.cell(row=row_num, column=10).value = vfm_cat
                ws.cell(row=row_num, column=8).value = vfm_cat

            else:
                vfm_cat = current_data[project_name]["VfM Category single entry"]
                # ws.cell(row=row_num, column=10).value = vfm_cat
                ws.cell(row=row_num, column=8).value = vfm_cat

            # """vfm category baseline"""
            # bl_i = master.bl_index["ipdc_benefits"][project_name][2]
            # try:
            #     if (
            #             master.master_data[bl_i]["data"][project_name][
            #                 "VfM Category single entry"
            #             ]
            #             is None
            #     ):
            #         vfm_cat_baseline = (
            #                 str(
            #                     master.master_data[bl_i]["data"][project_name][
            #                         "VfM Category lower range"
            #                     ]
            #                 )
            #                 + " - "
            #                 + str(
            #             master.master_data[bl_i]["data"][project_name][
            #                 "VfM Category upper range"
            #             ]
            #         )
            #         )
            #         # ws.cell(row=row_num, column=11).value = vfm_cat_baseline
            #     else:
            #         vfm_cat_baseline = master.master_data[bl_i]["data"][project_name][
            #             "VfM Category single entry"
            #         ]
            #         # ws.cell(row=row_num, column=11).value = vfm_cat_baseline
            #
            # except KeyError:
            #     try:
            #         vfm_cat_baseline = master.master_data[bl_i]["data"][project_name][
            #             "VfM Category single entry"
            #         ]
            #         # ws.cell(row=row_num, column=11).value = vfm_cat_baseline
            #     except KeyError:
            #         vfm_cat_baseline = master.master_data[bl_i]["data"][project_name][
            #             "VfM Category"
            #         ]
            #         # ws.cell(row=row_num, column=11).value = vfm_cat_baseline
            #
            # if vfm_cat != vfm_cat_baseline:
            #     if vfm_cat_baseline is None:
            #         pass
            #     else:
            #         ws.cell(row=row_num, column=8).font = Font(
            #             name="Arial", size=8, color="00fc2525"
            #         )

            abb = master.abbreviations[project_name]["abb"]
            current = get_milestone_date(
                abb,
                milestones.milestone_dict,
                str(master.current_quarter),
                "Full Operations",
            )
            # if current == None:
            #     current = "None"
            last_quarter = get_milestone_date(
                abb,
                milestones.milestone_dict,
                str(master.quarter_list[1]),
                "Full Operations",
            )
            # bl = get_milestone_date(
            #     abb, milestones.milestone_dict, "bl_one", " Full Operations"
            # )
            ws.cell(row=row_num, column=9).value = current
            if current is not None and current < IPDC_DATE:
                ws.cell(row=row_num, column=9).value = "Completed"
            if current is None:
                ws.cell(row=row_num, column=9).value = "None"

            try:
                last_change = (current - last_quarter).days
                ws.cell(row=row_num, column=10).value = plus_minus_days(last_change)
            except TypeError:
                pass
            # try:
            #     bl_change = (current - bl).days
            #     ws.cell(row=row_num, column=11).value = plus_minus_days(bl_change)
            #     if bl_change is not None and bl_change > 85:
            #         ws.cell(row=row_num, column=11).font = Font(
            #             name="Arial", size=10, color="00fc2525"
            #         )
            # except TypeError:
            #     pass

            # last at/next at ipdc information  removed
            # try:
            #     ws.cell(row=row_num, column=12).value = concatenate_dates(
            #         master.master_data[0].data[project_name]["Last time at BICC"],
            #         IPDC_DATE,
            #     )
            #     ws.cell(row=row_num, column=13).value = concatenate_dates(
            #         master.master_data[0].data[project_name]["Next at BICC"],
            #         IPDC_DATE,
            #     )
            # except (KeyError, TypeError):
            #     print(
            #         project_name
            #         + " last at / next at ipdc data could not be calculated. Check data."
            #     )

            """IPA DCA rating"""
            try:
                ipa_dca = CONVERT_RAG(current_data[project_name]["GMPP - IPA DCA"])
            except KeyError:
                raise InputError(
                    "No GMPP IPA DCA key in quarter master. This key must"
                    " be present for dashboard compilation. Stopping."
                )
            ws.cell(row=row_num, column=15).value = ipa_dca
            if ipa_dca == "None":
                ws.cell(row=row_num, column=15).value = ""

            # SRO forward look
            try:
                fwd_look = current_data[project_name]["SRO Forward Look Assessment"]
            except KeyError:
                raise InputError(
                    "No SRO Forward Look Assessment key in current quarter master. This key must"
                    " be present for dashboard compilation. Stopping."
                )
            if fwd_look == "Worsening":
                ws.cell(row=row_num, column=18).value = 1
            if fwd_look == "No Change Expected":
                ws.cell(row=row_num, column=18).value = 2
            if fwd_look == "Improving":
                ws.cell(row=row_num, column=18).value = 3
            if fwd_look is None:
                ws.cell(row=row_num, column=18).value = ""

            """SRO three DCA rating"""
            sro_dca_three = CONVERT_RAG(
                current_data[project_name]["Departmental DCA"]
            )  # "GMPP - SRO DCA"
            ws.cell(row=row_num, column=16).value = sro_dca_three
            if sro_dca_three == "None":
                ws.cell(row=row_num, column=16).value = ""

            """DCA rating - this quarter"""
            ws.cell(row=row_num, column=19).value = CONVERT_RAG(
                current_data[project_name]["Departmental DCA"]
            )
            """DCA rating - last qrt"""
            try:
                ws.cell(row=row_num, column=20).value = CONVERT_RAG(
                    last_data[project_name]["Departmental DCA"]
                )
            except KeyError:
                ws.cell(row=row_num, column=20).value = ""
            """DCA rating - 2 qrts ago"""
            try:
                ws.cell(row=row_num, column=21).value = CONVERT_RAG(
                    master.master_data[2]["data"][project_name]["Departmental DCA"]
                )
            except (KeyError, IndexError):
                ws.cell(row=row_num, column=21).value = ""
            """DCA rating - 3 qrts ago"""
            try:
                ws.cell(row=row_num, column=22).value = CONVERT_RAG(
                    master.master_data[3]["data"][project_name]["Departmental DCA"]
                )
            except (KeyError, IndexError):
                ws.cell(row=row_num, column=22).value = ""

            # """DCA rating - baseline"""
            # bl_i = master.bl_index["ipdc_costs"][project_name][2]
            # ws.cell(row=row_num, column=23).value = CONVERT_RAG(
            #     master.master_data[bl_i]["data"][project_name]["Departmental DCA"]
            # )

    # places arrow icons for sro forward look assessment
    icon_set_rule = IconSetRule("3Arrows", "num", ["1", "2", "3"], showValue=False)
    ws.conditional_formatting.add("R4:R40", icon_set_rule)

    """list of columns with conditional formatting"""
    list_columns = ["o", "s", "t", "u", "v", "w"]

    """same loop but the text is black. In addition these two loops go through the list_columns list above"""
    for column in list_columns:
        for i, dca in enumerate(rag_txt_list):
            text = black_text
            fill = fill_colour_list[i]
            dxf = DifferentialStyle(font=text, fill=fill)
            rule = Rule(type="containsText", operator="containsText", text=dca, dxf=dxf)
            for_rule_formula = 'NOT(ISERROR(SEARCH("' + dca + '",' + column + "5)))"
            rule.formula = [for_rule_formula]
            ws.conditional_formatting.add(column + "5:" + column + "60", rule)

    for row_num in range(2, ws.max_row + 1):
        for col_num in range(5, ws.max_column + 1):
            if ws.cell(row=row_num, column=col_num).value == 0:
                ws.cell(row=row_num, column=col_num).value = "-"

    return wb
