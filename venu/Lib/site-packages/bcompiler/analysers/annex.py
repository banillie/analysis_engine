import datetime
import os
import sys

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.styles.fills import PatternFill

from .utils import MASTER_XLSX, logger, get_number_of_projects, project_titles_in_master
from ..utils import ROOT_PATH, runtime_config, CONFIG_FILE, project_data_from_master

runtime_config.read(CONFIG_FILE)

# Fill colours
red_colour = Color(rgb='00fc2525')
red_amber_colour = Color(rgb='00f97b31')
amber_colour = Color(rgb='00fce553')
amber_green_colour = Color(rgb='00a5b700')
green_colour = Color(rgb='0017960c')

red = PatternFill(patternType='solid', fgColor=red_colour, bgColor="000000")
red_amber = PatternFill(patternType='solid', fgColor=red_amber_colour, bgColor="000000")
amber = PatternFill(patternType='solid', fgColor=amber_colour, bgColor="000000")
amber_green = PatternFill(patternType='solid', fgColor=amber_green_colour, bgColor="000000")
green = PatternFill(patternType='solid', fgColor=green_colour, bgColor="000000")

fg_colour_map = {
    "Red": red,
    "Amber/Red": red_amber,
    "Amber": amber,
    "Amber/Green": amber_green,
    "Green": green
}


def abbreviate_project_stage(stage: str):
    if stage == "Outline Business Case":
        return "OBC"
    elif stage == "Strategic Outline Case" or stage == "Strategic Outline Business Case":
        return "SOBC"
    elif stage == "Full Business Case":
        return "FBC"
    else:
        return "UNKNOWN STAGE"


def process_master(source_wb, project_number, dca_map, diff: list):
    """
    Function which is called on each cycle in main loop. Takes a master workbook
    and a project number as arguments. Creates a new workbook, populates it with
     the required data from the source_wb file passed in, formats it, then returns
     the workbook from the function, along with the project name which is used
     to name the file. d_map is a dict of DCA values for each project
    """

    wb = Workbook()
    sheet = wb.active
    ws2 = source_wb.active

    al = Alignment(horizontal="left", vertical="top", wrap_text=True,
                   shrink_to_fit=True)

    al2 = Alignment(horizontal="center", vertical="center", wrap_text=True,
                    shrink_to_fit=True)

    al_right = Alignment(horizontal="right", vertical="bottom", wrap_text=True,
                         shrink_to_fit=True)

    double_bottom_border = Border(left=Side(style='none'),
                                  right=Side(style='none'),
                                  top=Side(style='none'),
                                  bottom=Side(style='double'))

    single_bottom_border = Border(left=Side(style='none'),
                                  right=Side(style='none'),
                                  top=Side(style='none'),
                                  bottom=Side(style='thick'))

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    bold18Font = Font(size=18, bold=True)

    project_name = ws2.cell(row=1, column=project_number).value
    SRO_name = ws2.cell(row=59, column=project_number).value
    WLC_value = ws2.cell(row=304, column=project_number).value
    project_stage = abbreviate_project_stage(ws2.cell(row=281, column=project_number).value)
    SRO_conf = ws2.cell(row=57, column=project_number).value
    # SRO_conf_last_qtr =
    SoP = ws2.cell(row=201, column=project_number).value
    ipa_rag = ws2.cell(row=1274, column=project_number).value
    if isinstance(SoP, datetime.datetime):
        SoP = SoP.date()
    finance_DCA = ws2.cell(row=280, column=project_number).value
    benefits_DCA = ws2.cell(row=1152, column=project_number).value
    SRO_Comm = ws2.cell(row=58, column=project_number).value
# red_color = 'ffc7ce'
# red_fill = styles.PatternFill(start_color=red_color, end_color=red_color, fill_type='solid')
# sheet.conditional_formatting.add('B5', CellIsRule(operator='containsText', formula=['Amber/Green'], fill=red_fill))
    sheet.column_dimensions['A'].width = 15
    sheet.column_dimensions['B'].width = 15
    sheet.column_dimensions['C'].width = 15
    sheet.column_dimensions['D'].width = 15
    sheet.column_dimensions['E'].width = 15
    sheet.column_dimensions['F'].width = 15
    sheet.row_dimensions[1].height = 30
    sheet.merge_cells('A1:F1')
    sheet['A1'].value = project_name
    sheet['A1'].font = bold18Font
    sheet['A1'].border = double_bottom_border
    sheet['B1'].border = double_bottom_border
    sheet['C1'].border = double_bottom_border
    sheet['D1'].border = double_bottom_border
    sheet['E1'].border = double_bottom_border
    sheet['F1'].border = double_bottom_border
    sheet['A1'].alignment = al2
    sheet.row_dimensions[2].height = 10
    sheet.row_dimensions[3].height = 20
    sheet['A3'].value = 'SRO'
    sheet.merge_cells('C3:F3')
    sheet['C3'].value = SRO_name
    sheet.row_dimensions[4].height = 10
    sheet.row_dimensions[5].height = 20
    sheet['A5'].value = 'WLC(£m):'
    sheet['B5'].value = float("{:.1f}".format(WLC_value))
    sheet['B5'].alignment = al_right
    sheet['C5'].value = 'Project Stage:'
    sheet['D5'].value = project_stage
    sheet['E5'].value = 'Start of Ops:'
    sheet['F5'].value = SoP
    sheet.row_dimensions[6].height = 10
    sheet.row_dimensions[7].height = 20
    sheet['A7'].value = 'DCA now'
    sheet['B7'].value = SRO_conf
    sheet['B7'].border = thin_border
    sheet['C7'].value = 'DCA last quarter'
    if not project_name in diff:
        sheet['D7'].value = dca_map[project_name]
    sheet['D7'].border = thin_border
    sheet['E7'].value = 'IPA DCA'
    sheet['F7'].border = thin_border
    sheet['F7'].value = ipa_rag
    sheet.row_dimensions[8].height = 10
    sheet.row_dimensions[9].height = 20
    sheet['A9'].value = 'Finance DCA'
    sheet['B9'].value = finance_DCA
    sheet['B9'].border = thin_border
    sheet['C9'].value = 'Benefits DCA'
    sheet['D9'].value = benefits_DCA
    sheet['D9'].border = thin_border
    sheet.row_dimensions[10].height = 10
    sheet['A11'].value = SRO_Comm
    sheet['A11'].alignment = al
    sheet['B11'].border = double_bottom_border
    sheet['C11'].border = double_bottom_border
    sheet['D11'].border = double_bottom_border
    sheet['E11'].border = double_bottom_border
    sheet['F11'].border = double_bottom_border
    sheet['A40'].border = double_bottom_border
    sheet['B40'].border = double_bottom_border
    sheet['C40'].border = double_bottom_border
    sheet['D40'].border = double_bottom_border
    sheet['E40'].border = double_bottom_border
    sheet['F40'].border = double_bottom_border
    sheet.merge_cells('A11:F45')
    sheet['A45'].border = single_bottom_border

    # set print area
    sheet.print_area = "A1:F45"


    def _pattern(str_colour: str):
        return fg_colour_map[str_colour]

    for row in sheet.iter_rows(min_row=7, max_col=6, max_row=9):
        for cell in row:
            if cell.value in ['Green', 'Amber/Green', 'Amber', 'Amber/Red', 'Red']:
                cell.fill = _pattern(cell.value)

    return wb, project_name  # outputs a tuple of (wb, project_name) <- parens are optional!


def _dca_map(master_file: str):
    d = project_data_from_master(master_file)
    ds = {}
    for item in d.items():
        ds.update({item[0]: item[1]['Departmental DCA']})
    return ds



def run(compare_master=None, output_path=None, user_provided_master_path=None):

    if user_provided_master_path:
        logger.info(f"Using master file: {user_provided_master_path}")
        q2 = load_workbook(user_provided_master_path)
        projects_in_current_master = project_titles_in_master(user_provided_master_path)
    else:
        logger.info(f"Using default master file (refer to config.ini)")
        q2 = load_workbook(MASTER_XLSX)
        projects_in_current_master = project_titles_in_master(MASTER_XLSX)

    if compare_master:
        projects_in_compare_master = project_titles_in_master(compare_master)
        diff = set(projects_in_current_master).difference(projects_in_compare_master)
        if diff:
            logger.warning("{} not present in compare master.".format(", ".join(diff)))
        dca_map = _dca_map(compare_master)
        logger.info(f"Running annex analyser using {compare_master} as comparison.")
    else:
        compare_master = os.path.join(
            ROOT_PATH, runtime_config['AnalyserAnnex']['compare_master'])
        projects_in_compare_master = project_titles_in_master(compare_master)
        diff = set(projects_in_current_master).difference(projects_in_compare_master)
        if diff:
            logger.warning("{} not present in compare master.".format(", ".join(diff)))
        try:
            dca_map = _dca_map(compare_master)
        except FileNotFoundError:
            logger.critical(f"Cannot find {compare_master} in /Documents/bcompiler directory. Either put it there or"
                            f" or call annex with --compare option.")
            sys.exit(1)
        logger.info(f"Running annex analyser using {compare_master} as comparison.")


    # get the number of projects, so we know how many times to loop
    project_count = get_number_of_projects(q2)

    for p in range(2, project_count + 2):  # start at 2, representating col B in master; go until number of projects plus 2

        # pass out master and project number into the process_master() function
        # we capture the workbook object and project name in a tuple (these are the objects passed out by the return statement inside process_master() function
        output_wb, project_name = process_master(q2, p, dca_map, diff)
        if '/' in project_name:
            project_name = project_name.replace('/', '_')

        # save the file, using the project_name variable in the file name
        try:
            if output_path:
                output_wb.save(os.path.join(output_path[0], f'{project_name}_ANNEX.xlsx'))
                logger.info(f"{project_name}_ANNEX.xlsx to {output_path}")
            else:
                output_path = os.path.join(ROOT_PATH, 'output')
                output_wb.save(os.path.join(output_path, f'{project_name}_ANNEX.xlsx'))
                logger.info(f"{project_name}_ANNEX.xlsx to {output_path}")
                output_path = ""
        except PermissionError:
            logger.critical(f"Cannot save {project_name}_ANNEX.xlsx file - you already have it open. Close and run again.")
            return


if __name__ == "__main__":
    run()
