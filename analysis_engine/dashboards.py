from openpyxl import Workbook
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting import Rule

from analysis_engine.dictionaries import (
    DATA_KEY_DICT,
    DASHBOARD_BC_STAGE_ABBREVIATION,
    CONVERT_RAG,
    rag_txt_list,
    conf_list,
    risk_list,
)
from analysis_engine.dandelion import dandelion_number_text
from analysis_engine.colouring import black_text, fill_colour_list


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
            ws.cell(row=row_num, column=5).value = DASHBOARD_BC_STAGE_ABBREVIATION[
                bc_stage
            ]
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
            ws.cell(row=row_num, column=5).value = DASHBOARD_BC_STAGE_ABBREVIATION[
                bc_stage
            ]
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
                dca = CONVERT_RAG[
                    master["master_data"][0]["data"][project_name][key]
                ]
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
