import configparser
import csv
import io
import os
import re
import shutil
import tempfile
from datetime import date, datetime

import pytest
from openpyxl import load_workbook, Workbook

from ..utils import generate_test_template_from_real as gen_template

TEMPDIR = tempfile.gettempdir()

AUX_DIR = "/".join([TEMPDIR, 'bcompiler'])
SOURCE_DIR = "/".join([AUX_DIR, 'source'])
RETURNS_DIR = "/".join([SOURCE_DIR, 'returns'])
OUTPUT_DIR = "/".join([AUX_DIR, 'output'])

try:
    os.mkdir(AUX_DIR)
except (FileExistsError, IsADirectoryError):
    shutil.rmtree(AUX_DIR)
    os.mkdir(AUX_DIR)
    os.mkdir(OUTPUT_DIR)
    os.mkdir(SOURCE_DIR)
    os.mkdir(RETURNS_DIR)

config = configparser.ConfigParser()
CONFIG_FILE = 'test_config.ini'
config.read(CONFIG_FILE)

BICC_TEMPLATE_FOR_TESTS = config['Template']['ActualTemplatePath']

datamap_header = "cell_key,template_sheet,cell_reference,verification"

real_datamap_data = """
Project/Programme Name,Summary,B5,
SRO Sign-Off,Summary,B49,
Reporting period (GMPP - Snapshot Date),Summary,G3,
Quarter Joined,Summary,I3,
GMPP (GMPP - formally joined GMPP),Summary,G5,
IUK top 40,Summary,G6,
Top 37,Summary,I5,
DfT Business Plan,Summary,I6,
DFT ID Number,Summary,B6,
MPA ID Number,Summary,C6,
Working Contact Name,Summary,H8,
Working Contact Telephone,Summary,H9,
Working Contact Email,Summary,H10,
DfT Group,Summary,B8,DfT Group,
DfT Division,Summary,B9,DfT Division,
Agency or delivery partner (GMPP - Delivery Organisation primary),Summary,B10,Agency,
Strategic Alignment/Government Policy (GMPP - Key drivers),Summary,B26,
Project Scope,Summary,G26,
Brief project description (GMPP - brief descripton),Summary,G27,
Delivery Structure,Summary,G32,Entity format,
Description if 'Other',Summary,G33,
Change Delivery Methodology,Summary,G39,Methodology,
Primary Category,Summary,H41,Category,
If other please describe,Summary,J41,
Secondary Category,Summary,H42,Category,
If other please describe,Summary,J42,
Tertiary Category,Summary,H43,Category,
If other please describe,Summary,J43,
Has Project Scope Changed?,Summary,G35,Scope Changed,
Scope Change Commentary (if applicable),Summary,H35,
List Strategic Outcomes (GMPP - Intended Outcome 1),Summary,B33,
IO1 - Monetised?,Summary,D33,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 2),Summary,B34,
IO2 - Monetised?,Summary,D34,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 3),Summary,B35,
IO3 - Monetised?,Summary,D35,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 4),Summary,B36,
IO4 - Monetised?,Summary,D36,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 5),Summary,B37,
IO5 - Monetised?,Summary,D37,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 6),Summary,B38,
IO6 - Monetised?,Summary,D38,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 7),Summary,B39,
IO7 - Monetised?,Summary,D39,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 8),Summary,B40,
IO8 - Monetised?,Summary,D40,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 9),Summary,B41,
IO9 - Monetised?,Summary,D41,Monetised / Non Monetised Benefits,
List Strategic Outcomes (GMPP - Intended Outcome 10),Summary,B42,
IO10 - Monetised?,Summary,D42,Monetised / Non Monetised Benefits,
Single Departmental Plan 1,Summary,B28,SDP,
Single Departmental Plan 2,Summary,B29,SDP,
Single Departmental Plan 3,Summary,B30,SDP,
Single Departmental Plan 4,Summary,B31,SDP,
Risk Level (RPA),Summary,H29,RPA level,
Risk Level (RPA) Date,Summary,H30,
Departmental DCA,Summary,B46,RAG,
Departmental DCA Narrative,Summary,C47,
SRO Full Name,Summary,C12,
SRO Email,Summary,C14,
SRO Phone No.,Summary,C13,
Percentage of time spent on SRO Role,Summary,C18,Percentage of time spent on SRO role,
SRO Tenure Start Date,Summary,C15,
SRO Tenure End Date,Summary,C17,
SRO MPLA Status,Summary,C21,MPLA / PLP,
SRO MPLA - If 'other' please describe,Summary,C22,
SRO PLP Status,Summary,C23,MPLA / PLP,
SRO PLP - If 'other' please describe,Summary,C24,
Date If Current SRO Letter Issued,Summary,C16,
Has the SRO changed?,Summary,C19,Yes/No,
If new SRO reason for change,Summary,C20,PL Changes,
PD Full Name,Summary,H12,
PD Email,Summary,H14,
PD Phone No.,Summary,H13,
Secondary PDs please name?,Summary,J12,
PD MPLA Status,Summary,H21,MPLA / PLP,
PD MPLA - If 'other' please describe,Summary,H22,
PD PLP Status,Summary,H23,MPLA / PLP,
PD PLP - If 'other' please describe,Summary,H24,
Has PD changed?,Summary,H19,Yes/No,
If new PD reason for change,Summary,H20,PL Changes,
Percentage of time spent on PD Role,Summary,H18,Percentage of time spent on SRO role,
PD Tenure Start Date,Summary,H15,
PD Tenure End Date,Summary,H17,
Date if PD letter issued,Summary,H16,
Project stage,Approval & Project milestones,B5,Project stage,
Project stage if Other,Approval & Project milestones,D5,
Last time at BICC,Approval & Project milestones,B4,
Next at BICC,Approval & Project milestones,D4,
Approval MM1,Approval & Project milestones,A9,
Approval MM1 Original Baseline,Approval & Project milestones,B9,
Approval MM1 Latest Approved Baseline,Approval & Project milestones,C9,
Approval MM1 Forecast / Actual,Approval & Project milestones,D9,
Approval MM1 Milestone Type,Approval & Project milestones,E9,Milestone Types,
Approval MM1 Notes,Approval & Project milestones,F9,
Approval MM2,Approval & Project milestones,A10,
Approval MM2 Original Baseline,Approval & Project milestones,B10,
Approval MM2 Latest Approved Baseline,Approval & Project milestones,C10,
Approval MM2 Forecast / Actual,Approval & Project milestones,D10,
Approval MM2 Milestone Type,Approval & Project milestones,E10,
Approval MM2 Notes,Approval & Project milestones,F10,
Approval MM3,Approval & Project milestones,A11,
Approval MM3 Original Baseline,Approval & Project milestones,B11,
Approval MM3 Latest Approved Baseline,Approval & Project milestones,C11,
Approval MM3 Forecast / Actual,Approval & Project milestones,D11,
Approval MM3 Milestone Type,Approval & Project milestones,E11,Milestone Types,
Approval MM3 Notes,Approval & Project milestones,F11,
Approval MM4,Approval & Project milestones,A12,
Approval MM4 Original Baseline,Approval & Project milestones,B12,
Approval MM4 Latest Approved Baseline,Approval & Project milestones,C12,
Approval MM4 Forecast / Actual,Approval & Project milestones,D12,
Approval MM4 Milestone Type,Approval & Project milestones,E12,
Approval MM4 Notes,Approval & Project milestones,F12,
Approval MM5,Approval & Project milestones,A13,
Approval MM5 Original Baseline,Approval & Project milestones,B13,
Approval MM5 Latest Approved Baseline,Approval & Project milestones,C13,
Approval MM5 Forecast / Actual,Approval & Project milestones,D13,
Approval MM5 Milestone Type,Approval & Project milestones,E13,Milestone Types,
Approval MM5 Notes,Approval & Project milestones,F13,
Approval MM6,Approval & Project milestones,A14,
Approval MM6 Original Baseline,Approval & Project milestones,B14,
Approval MM6 Latest Approved Baseline,Approval & Project milestones,C14,
Approval MM6 Forecast / Actual,Approval & Project milestones,D14,
Approval MM6 Milestone Type,Approval & Project milestones,E14,Milestone Types,
Approval MM6 Notes,Approval & Project milestones,F14,
Approval MM7,Approval & Project milestones,A15,
Approval MM7 Original Baseline,Approval & Project milestones,B15,
Approval MM7 Latest Approved Baseline,Approval & Project milestones,C15,
Approval MM7 Forecast / Actual,Approval & Project milestones,D15,
Approval MM7 Milestone Type,Approval & Project milestones,E15,Milestone Types,
Approval MM7 Notes,Approval & Project milestones,F15,
Approval MM8,Approval & Project milestones,A16,
Approval MM8 Original Baseline,Approval & Project milestones,B16,
Approval MM8 Latest Approved Baseline,Approval & Project milestones,C16,
Approval MM8 Forecast / Actual,Approval & Project milestones,D16,
Approval MM8 Milestone Type,Approval & Project milestones,E16,Milestone Types,
Approval MM8 Notes,Approval & Project milestones,F16,
Approval MM9,Approval & Project milestones,A17,
Approval MM9 Original Baseline,Approval & Project milestones,B17,
Approval MM9 Latest Approved Baseline,Approval & Project milestones,C17,
Approval MM9 Forecast / Actual,Approval & Project milestones,D17,
Approval MM9 Milestone Type,Approval & Project milestones,E17,Milestone Types,
Approval MM9 Notes,Approval & Project milestones,F17,
Approval MM10,Approval & Project milestones,A18,
Approval MM10 Original Baseline,Approval & Project milestones,B18,
Approval MM10 Latest Approved Baseline,Approval & Project milestones,C18,
Approval MM10 Forecast / Actual,Approval & Project milestones,D18,
Approval MM10 Milestone Type,Approval & Project milestones,E18,Milestone Types,
Approval MM10 Notes,Approval & Project milestones,F18,
Approval MM11,Approval & Project milestones,A19,
Approval MM11 Original Baseline,Approval & Project milestones,B19,
Approval MM11 Latest Approved Baseline,Approval & Project milestones,C19,
Approval MM11 Forecast / Actual,Approval & Project milestones,D19,
Approval MM11 Milestone Type,Approval & Project milestones,E19,
Approval MM11 Notes,Approval & Project milestones,F19,
Approval MM12,Approval & Project milestones,A20,
Approval MM12 Original Baseline,Approval & Project milestones,B20,
Approval MM12 Latest Approved Baseline,Approval & Project milestones,C20,
Approval MM12 Forecast / Actual,Approval & Project milestones,D20,
Approval MM12 Milestone Type,Approval & Project milestones,E20,Milestone Types,
Approval MM12 Notes,Approval & Project milestones,F20,
Approval MM13,Approval & Project milestones,A21,
Approval MM13 Original Baseline,Approval & Project milestones,B21,
Approval MM13 Latest Approved Baseline,Approval & Project milestones,C21,
Approval MM13 Forecast - Actual,Approval & Project milestones,D21,
Approval MM13 Type,Approval & Project milestones,E21,Milestone Types,
Approval MM13 Notes,Approval & Project milestones,F21,
Approval MM14,Approval & Project milestones,A22,
Approval MM14 Original Baseline,Approval & Project milestones,B22,
Approval MM14 Latest Approved Baseline,Approval & Project milestones,C22,
Approval MM14 Forecast - Actual,Approval & Project milestones,D22,
Approval MM14 Type,Approval & Project milestones,E22,Milestone Types,
Approval MM14 Notes,Approval & Project milestones,F22,
Approval MM15,Approval & Project milestones,A23,
Approval MM15 Original Baseline,Approval & Project milestones,B23,
Approval MM15 Latest Approved Baseline,Approval & Project milestones,C23,
Approval MM15 Forecast - Actual,Approval & Project milestones,D23,
Approval MM15 Type,Approval & Project milestones,E23,Milestone Types,
Approval MM15 Notes,Approval & Project milestones,F23,
Approval MM16,Approval & Project milestones,A24,
Approval MM16 Original Baseline,Approval & Project milestones,B24,
Approval MM16 Latest Approved Baseline,Approval & Project milestones,C24,
Approval MM16 Forecast - Actual,Approval & Project milestones,D24,
Approval MM16 Type,Approval & Project milestones,E24,Milestone Types,
Approval MM16 Notes,Approval & Project milestones,F24,
Project MM18,Approval & Project milestones,A26,
Project MM18 Original Baseline,Approval & Project milestones,B26,
Project MM18 Latest Approved Baseline,Approval & Project milestones,C26,
Project MM18 Forecast - Actual,Approval & Project milestones,D26,
Project MM18 Type,Approval & Project milestones,E26,
Project MM18 Notes,Approval & Project milestones,F26,
Project MM19,Approval & Project milestones,A27,
Project MM19 Original Baseline,Approval & Project milestones,B27,
Project MM19 Latest Approved Baseline,Approval & Project milestones,C27,
Project MM19 Forecast - Actual,Approval & Project milestones,D27,
Project MM19 Type,Approval & Project milestones,E27,Milestone Types,
Project MM19 Notes,Approval & Project milestones,F27,
Project MM20,Approval & Project milestones,A28,
Project MM20 Original Baseline,Approval & Project milestones,B28,
Project MM20 Latest Approved Baseline,Approval & Project milestones,C28,
Project MM20 Forecast - Actual,Approval & Project milestones,D28,
Project MM20 Type,Approval & Project milestones,E28,Milestone Types,
Project MM20 Notes,Approval & Project milestones,F28,
Project MM21,Approval & Project milestones,A29,
Project MM21 Original Baseline,Approval & Project milestones,B29,
Project MM21 Latest Approved Baseline,Approval & Project milestones,C29,
Project MM21 Forecast - Actual,Approval & Project milestones,D29,
Project MM21 Type,Approval & Project milestones,E29,
Project MM21 Notes,Approval & Project milestones,F29,
Project MM22,Approval & Project milestones,A30,
Project MM22 Original Baseline,Approval & Project milestones,B30,
Project MM22 Latest Approved Baseline,Approval & Project milestones,C30,
Project MM22 Forecast - Actual,Approval & Project milestones,D30,
Project MM22 Type,Approval & Project milestones,E30,
Project MM22 Notes,Approval & Project milestones,F30,
Project MM23,Approval & Project milestones,A31,
Project MM23 Original Baseline,Approval & Project milestones,B31,
Project MM23 Latest Approved Baseline,Approval & Project milestones,C31,
Project MM23 Forecast - Actual,Approval & Project milestones,D31,
Project MM23 Type,Approval & Project milestones,E31,Milestone Types,
Project MM23 Notes,Approval & Project milestones,F31,
Project MM24,Approval & Project milestones,A32,
Project MM24 Original Baseline,Approval & Project milestones,B32,
Project MM24 Latest Approved Baseline,Approval & Project milestones,C32,
Project MM24 Forecast - Actual,Approval & Project milestones,D32,
Project MM24 Type,Approval & Project milestones,E32,Milestone Types,
Project MM24 Notes,Approval & Project milestones,F32,
Project MM25,Approval & Project milestones,A33,
Project MM25 Original Baseline,Approval & Project milestones,B33,
Project MM25 Latest Approved Baseline,Approval & Project milestones,C33,
Project MM25 Forecast - Actual,Approval & Project milestones,D33,
Project MM25 Type,Approval & Project milestones,E33,Milestone Types,
Project MM25 Notes,Approval & Project milestones,F33,
Project MM26,Approval & Project milestones,A34,
Project MM26 Original Baseline,Approval & Project milestones,B34,
Project MM26 Latest Approved Baseline,Approval & Project milestones,C34,
Project MM26 Forecast - Actual,Approval & Project milestones,D34,
Project MM26 Type,Approval & Project milestones,E34,Milestone Types,
Project MM26 Notes,Approval & Project milestones,F34,
Project MM27,Approval & Project milestones,A35,
Project MM27 Original Baseline,Approval & Project milestones,B35,
Project MM27 Latest Approved Baseline,Approval & Project milestones,C35,
Project MM27 Forecast - Actual,Approval & Project milestones,D35,
Project MM27 Type,Approval & Project milestones,E35,Milestone Types,
Project MM27 Notes,Approval & Project milestones,F35,
Project MM28,Approval & Project milestones,A36,
Project MM28 Original Baseline,Approval & Project milestones,B36,
Project MM28 Latest Approved Baseline,Approval & Project milestones,C36,
Project MM28 Forecast - Actual,Approval & Project milestones,D36,
Project MM28 Type,Approval & Project milestones,E36,Milestone Types,
Project MM28 Notes,Approval & Project milestones,F36,
Project MM29,Approval & Project milestones,A37,
Project MM29 Original Baseline,Approval & Project milestones,B37,
Project MM29 Latest Approved Baseline,Approval & Project milestones,C37,
Project MM29 Forecast - Actual,Approval & Project milestones,D37,
Project MM29 Type,Approval & Project milestones,E37,Milestone Types,
Project MM29 Notes,Approval & Project milestones,F37,
Project MM30,Approval & Project milestones,A38,
Project MM30 Original Baseline,Approval & Project milestones,B38,
Project MM30 Latest Approved Baseline,Approval & Project milestones,C38,
Project MM30 Forecast - Actual,Approval & Project milestones,D38,
Project MM30 Type,Approval & Project milestones,E38,Milestone Types,
Project MM30 Notes,Approval & Project milestones,F38,
Project MM31,Approval & Project milestones,A39,
Project MM31 Original Baseline,Approval & Project milestones,B39,
Project MM31 Latest Approved Baseline,Approval & Project milestones,C39,
Project MM31 Forecast - Actual,Approval & Project milestones,D39,
Project MM31 Type,Approval & Project milestones,E39,Milestone Types,
Project MM31 Notes,Approval & Project milestones,F39,
Project MM32,Approval & Project milestones,A40,
Project MM32 Original Baseline,Approval & Project milestones,B40,
Project MM32 Latest Approved Baseline,Approval & Project milestones,C40,
Project MM32 Forecast - Actual,Approval & Project milestones,D40,
Project MM32 Type,Approval & Project milestones,E40,Milestone Types,
Project MM32 Notes,Approval & Project milestones,F40,
Milestone Commentary,Approval & Project milestones,B42,
Project Lifecycle Stage,Approval & Project milestones,B5,
Project Stage if Other,Approval & Project milestones,D5,
Significant Steel Requirement,Finance & Benefits,D15,Yes/No,
SRO Finance confidence,Finance & Benefits,C6,RAG 2,
BICC approval point,Finance & Benefits,E9,Business Cases,
Latest Treasury Approval Point (TAP) or equivalent,Finance & Benefits,E10,Business Cases,
Business Case used to source figures (GMPP TAP used to source figures),Finance & Benefits,C9,Business Cases,
Date of TAP used to source figures,Finance & Benefits,E11,
Name of source in not Business Case (GMPP -If not TAP please specify equivalent document used),Finance & Benefits,C10,
If not TAP please specify date of equivalent document,Finance & Benefits,C11,
Version Number Of Document used to Source Figures (GMPP - TAP version Number),Finance & Benefits,C12,
Date document approved by SRO,Finance & Benefits,C13,
Real or Nominal - Baseline,Finance & Benefits,C18,Finance figures format,
Real or Nominal - Actual/Forecast,Finance & Benefits,E18,Finance figures format,
Index Year,Finance & Benefits,B19,Index Years,
Deflator,Finance & Benefits,B20,Finance type,
Source of Finance,Finance & Benefits,B21,Finance type,
Other Finance type Description,Finance & Benefits,D21,
NPV for all projects and NPV for programmes if available,Finance & Benefits,B22,
Project cost to closure,Finance & Benefits,B23,
RDEL Total Budget/BL,Finance & Benefits,C72,
CDEL Total Budget/BL,Finance & Benefits,C125,
Non-Gov Total Budget/BL,Finance & Benefits,C135,
Total Budget/BL,Finance & Benefits,C136,
RDEL Total Forecast,Finance & Benefits,D133,
CDEL Total Forecast,Finance & Benefits,D134,
Non-Gov Total Forecast,Finance & Benefits,D135,
Total Forecast,Finance & Benefits,D136,
RDEL Total Variance,Finance & Benefits,E133,
CDEL Total Variance,Finance & Benefits,E134,
Non Gov Total Variance,Finance & Benefits,E135,
Total Variance,Finance & Benefits,E136,
RDEL Total Budget/BL SR (20/21),Finance & Benefits,H133,
CDEL Total Budget/BL SR (20/21),Finance & Benefits,H134,
Non-Gov Total Budget/BL SR (20/21),Finance & Benefits,H135,
Total Budget SR (20/21),Finance & Benefits,H136,
RDEL Total Forecast SR (20/21),Finance & Benefits,I133,
CDEL Total Forecast SR (20/21),Finance & Benefits,I134,
Non-Gov Total Forecast SR (20/21),Finance & Benefits,I135,
Total Forecast SR (20/21),Finance & Benefits,I136,
Project Costs Narrative RDEL,Finance & Benefits,B77,
Project Costs Narrative CDEL,Finance & Benefits,B130,
Pre 14-15 RDEL BL one off new costs,Finance & Benefits,C27,
Pre 14-15 RDEL BL recurring new costs,Finance & Benefits,D27,
Pre 14-15 RDEL BL recurring old costs,Finance & Benefits,E27,
Pre 14-15 RDEL BL Total,Finance & Benefits,F27,
Pre 14-15 RDEL Actual one off new costs,Finance & Benefits,C28,
Pre 14-15 RDEL Actual recurring new costs,Finance & Benefits,D28,
Pre 14-15 RDEL Actual recurring old costs,Finance & Benefits,E28,
Pre 14-15 RDEL Actual Total,Finance & Benefits,F28,
14-15 RDEL BL one off new costs,Finance & Benefits,C29,
14-15 RDEL BL recurring new costs,Finance & Benefits,D29,
14-15 RDEL BL recurring old costs,Finance & Benefits,E29,
14-15 RDEL BL Total,Finance & Benefits,F29,
14-15 RDEL Actual one off new costs,Finance & Benefits,C30,
14-15 RDEL Actual recurring new costs,Finance & Benefits,D30,
14-15 RDEL Actual recurring old costs,Finance & Benefits,E30,
14-15 RDEL Actual Total,Finance & Benefits,F30,
15-16 RDEL BL one off new costs,Finance & Benefits,C31,
15-16 RDEL BL recurring new costs,Finance & Benefits,D31,
15-16 RDEL BL recurring old costs,Finance & Benefits,E31,
15-16 RDEL BL Total,Finance & Benefits,F31,
15-16 RDEL Actual one off new costs,Finance & Benefits,C32,
15-16 RDEL Actual recurring new costs,Finance & Benefits,D32,
15-16 RDEL Actual recurring old costs,Finance & Benefits,E32,
15-16 RDEL Actual Total,Finance & Benefits,F32,
16-17 RDEL BL one off new costs,Finance & Benefits,C33,
16-17 RDEL BL recurring new costs,Finance & Benefits,D33,
16-17 RDEL BL recurring old costs,Finance & Benefits,E33,
16-17 RDEL BL Total,Finance & Benefits,F33,
16-17 RDEL Actual one off new costs,Finance & Benefits,C34,
16-17 RDEL Actual recurring new costs,Finance & Benefits,D34,
16-17 RDEL Actual recurring old costs,Finance & Benefits,E34,
16-17 RDEL Actual Total,Finance & Benefits,F34,
Pre 17-18 RDEL BL one off new costs,Finance & Benefits,C35,
Pre 17-18 RDEL BL recurring new costs,Finance & Benefits,D35,
Pre 17-18 RDEL BL recurring old costs,Finance & Benefits,E35,
Pre 17-18 RDEL BL Total,Finance & Benefits,F35,
Pre 17-18 RDEL Actual one off new costs,Finance & Benefits,C36,
Pre 17-18 RDEL Actual recurring new costs,Finance & Benefits,D36,
Pre 17-18 RDEL Actual recurring old costs,Finance & Benefits,E36,
Pre 17-18 RDEL Actual Total,Finance & Benefits,F36,
RDEL one off new cost spend 17-18 on profile,Finance & Benefits,C37,
RDEL recurring new cost spend 17-18 on profile,Finance & Benefits,D37,
RDEL recurring old cost spend 17-18 on profile,Finance & Benefits,E37,
RDEL total spend 17-18 on profile,Finance & Benefits,F37,
17-18 RDEL BL one off new costs,Finance & Benefits,C38,
17-18 RDEL BL recurring new costs,Finance & Benefits,D38,
17-18 RDEL BL recurring old costs,Finance & Benefits,E38,
17-18 RDEL BL Total,Finance & Benefits,F38,
17-18 RDEL Forecast one off new costs,Finance & Benefits,C39,
17-18 RDEL Forecast recurring new costs,Finance & Benefits,D39,
17-18 RDEL Forecast recurring old costs,Finance & Benefits,E39,
17-18 RDEL Forecast Total,Finance & Benefits,F39,
Pre 18-19 RDEL BL one off new costs,Finance & Benefits,C40,
Pre 18-19 RDEL BL recurring new costs,Finance & Benefits,D40,
Pre 18-19 RDEL BL recurring old costs,Finance & Benefits,E40,
Pre 18-19 RDEL BL Total,Finance & Benefits,F40,
Pre 18-19 RDEL Forecast one off new costs,Finance & Benefits,C41,
Pre 18-19 RDEL Forecast recurring new costs,Finance & Benefits,D41,
Pre 18-19 RDEL Forecast recurring old costs,Finance & Benefits,E41,
Pre 18-19 RDEL Forecast Total,Finance & Benefits,F41,
18-19 RDEL BL one off new costs,Finance & Benefits,C42,
18-19 RDEL BL recurring new costs,Finance & Benefits,D42,
18-19 RDEL BL recurring old costs,Finance & Benefits,E42,
18-19 RDEL BL Total,Finance & Benefits,F42,
18-19 RDEL Forecast one off new costs,Finance & Benefits,C43,
18-19 RDEL Forecast recurring new costs,Finance & Benefits,D43,
18-19 RDEL Forecast recurring old costs,Finance & Benefits,E43,
18-19 RDEL Forecast Total,Finance & Benefits,F43,
Pre 19-20 RDEL BL one off new costs,Finance & Benefits,C44,
Pre 19-20 RDEL BL recurring new costs,Finance & Benefits,D44,
Pre 19-20 RDEL BL recurring old costs,Finance & Benefits,E44,
Pre 19-20 RDEL BL Total,Finance & Benefits,F44,
Pre 19-20 RDEL Forecast one off new costs,Finance & Benefits,C45,
Pre 19-20 RDEL Forecast recurring new costs,Finance & Benefits,D45,
Pre 19-20 RDEL Forecast recurring old costs,Finance & Benefits,E45,
Pre 19-20 RDEL Forecast Total,Finance & Benefits,F45,
19-20 RDEL BL one off new costs,Finance & Benefits,C46,
19-20 RDEL BL recurring new costs,Finance & Benefits,D46,
19-20 RDEL BL recurring old costs,Finance & Benefits,E46,
19-20 RDEL BL Total,Finance & Benefits,F46,
19-20 RDEL Forecast one off new costs,Finance & Benefits,C47,
19-20 RDEL Forecast recurring new costs,Finance & Benefits,D47,
19-20 RDEL Forecast recurring old costs,Finance & Benefits,E47,
19-20 RDEL Forecast Total,Finance & Benefits,F47,
Pre 20-21 RDEL BL one off new costs,Finance & Benefits,C48,
Pre 20-21 RDEL BL recurring new costs,Finance & Benefits,D48,
Pre 20-21 RDEL BL recurring old costs,Finance & Benefits,E48,
Pre 20-21 RDEL BL Total,Finance & Benefits,F48,
Pre 20-21 RDEL Forecast one off new costs,Finance & Benefits,C49,
Pre 20-21 RDEL Forecast recurring new costs,Finance & Benefits,D49,
Pre 20-21 RDEL Forecast recurring old costs,Finance & Benefits,E49,
Pre 20-21 RDEL Forecast Total,Finance & Benefits,F49,
20-21 RDEL BL one off new costs,Finance & Benefits,C50,
20-21 RDEL BL recurring new costs,Finance & Benefits,D50,
20-21 RDEL BL recurring old costs,Finance & Benefits,E50,
20-21 RDEL BL Total,Finance & Benefits,F50,
20-21 RDEL Forecast one off new costs,Finance & Benefits,C51,
20-21 RDEL Forecast recurring new costs,Finance & Benefits,D51,
20-21 RDEL Forecast recurring old costs,Finance & Benefits,E51,
20-21 RDEL Forecast Total,Finance & Benefits,F51,
Pre 21-22 RDEL BL one off new costs,Finance & Benefits,C52,
Pre 21-22 RDEL BL recurring new costs,Finance & Benefits,D52,
Pre 21-22 RDEL BL recurring old costs,Finance & Benefits,E52,
Pre 21-22 RDEL BL Total,Finance & Benefits,F52,
Pre 21-22 RDEL Forecast one off new costs,Finance & Benefits,C53,
Pre 21-22 RDEL Forecast recurring new costs,Finance & Benefits,D53,
Pre 21-22 RDEL Forecast recurring old costs,Finance & Benefits,E53,
Pre 21-22 RDEL Forecast Total,Finance & Benefits,F53,
21-22 RDEL BL one off new costs,Finance & Benefits,C54,
21-22 RDEL BL recurring new costs,Finance & Benefits,D54,
21-22 RDEL BL recurring old costs,Finance & Benefits,E54,
21-22 RDEL BL Total,Finance & Benefits,F54,
21-22 RDEL Forecast one off new costs,Finance & Benefits,C55,
21-22 RDEL Forecast recurring new costs,Finance & Benefits,D55,
21-22 RDEL Forecast recurring old costs,Finance & Benefits,E55,
21-22 RDEL Forecast Total,Finance & Benefits,F55,
Pre 22-23 RDEL BL one off new costs,Finance & Benefits,C56,
Pre 22-23 RDEL BL recurring new costs,Finance & Benefits,D56,
Pre 22-23 RDEL BL recurring old costs,Finance & Benefits,E56,
Pre 22-23 RDEL BL Total,Finance & Benefits,F56,
Pre 22-23 RDEL Forecast one off new costs,Finance & Benefits,C57,
Pre 22-23 RDEL Forecast recurring new costs,Finance & Benefits,D57,
Pre 22-23 RDEL Forecast recurring old costs,Finance & Benefits,E57,
Pre 22-23 RDEL Forecast Total,Finance & Benefits,F57,
22-23 RDEL BL one off new costs,Finance & Benefits,C58,
22-23 RDEL BL recurring new costs,Finance & Benefits,D58,
22-23 RDEL BL recurring old costs,Finance & Benefits,E58,
22-23 RDEL BL Total,Finance & Benefits,F58,
22-23 RDEL Forecast one off new costs,Finance & Benefits,C59,
22-23 RDEL Forecast recurring new costs,Finance & Benefits,D59,
22-23 RDEL Forecast recurring old costs,Finance & Benefits,E59,
22-23 RDEL Forecast Total,Finance & Benefits,F59,
23-24 RDEL BL one off new costs,Finance & Benefits,C60,
23-24 RDEL BL recurring new costs,Finance & Benefits,D60,
23-24 RDEL BL recurring old costs,Finance & Benefits,E60,
23-24 RDEL BL Total,Finance & Benefits,F60,
23-24 RDEL Forecast one off new costs,Finance & Benefits,C61,
23-24 RDEL Forecast recurring new costs,Finance & Benefits,D61,
23-24 RDEL Forecast recurring old costs,Finance & Benefits,E61,
23-24 RDEL Forecast Total,Finance & Benefits,F61,
24-25 RDEL BL one off new costs,Finance & Benefits,C62,
24-25 RDEL BL recurring new costs,Finance & Benefits,D62,
24-25 RDEL BL recurring old costs,Finance & Benefits,E62,
24-25 RDEL BL Total,Finance & Benefits,F62,
24-25 RDEL Forecast one off new costs,Finance & Benefits,C63,
24-25 RDEL Forecast recurring new costs,Finance & Benefits,D63,
24-25 RDEL Forecast recurring old costs,Finance & Benefits,E63,
24-25 RDEL Forecast Total,Finance & Benefits,F63,
25-26 RDEL BL one off new costs,Finance & Benefits,C64,
25-26 RDEL BL recurring new costs,Finance & Benefits,D64,
25-26 RDEL BL recurring old costs,Finance & Benefits,E64,
25-26 RDEL BL Total,Finance & Benefits,F64,
25-26 RDEL Forecast one off new costs,Finance & Benefits,C65,
25-26 RDEL Forecast recurring new costs,Finance & Benefits,D65,
25-26 RDEL Forecast recurring old costs,Finance & Benefits,E65,
25-26 RDEL Forecast Total,Finance & Benefits,F65,
26-27 RDEL BL one off new costs,Finance & Benefits,C66,
26-27 RDEL BL recurring new costs,Finance & Benefits,D66,
26-27 RDEL BL recurring old costs,Finance & Benefits,E66,
26-27 RDEL BL Total,Finance & Benefits,F66,
26-27 RDEL Forecast one off new costs,Finance & Benefits,C67,
26-27 RDEL Forecast recurring new costs,Finance & Benefits,D67,
26-27 RDEL Forecast recurring old costs,Finance & Benefits,E67,
26-27 RDEL Forecast Total,Finance & Benefits,F67,
27-28 RDEL BL one off new costs,Finance & Benefits,C68,
27-28 RDEL BL recurring new costs,Finance & Benefits,D68,
27-28 RDEL BL recurring old costs,Finance & Benefits,E68,
27-28 RDEL BL Total,Finance & Benefits,F68,
27-28 RDEL Forecast one off new costs,Finance & Benefits,C69,
27-28 RDEL Forecast recurring new costs,Finance & Benefits,D69,
27-28 RDEL Forecast recurring old costs,Finance & Benefits,E69,
27-28 RDEL Forecast Total,Finance & Benefits,F69,
Unprofiled RDEL BL one off new costs,Finance & Benefits,C70,
Unprofiled RDEL BL recurring new costs,Finance & Benefits,D70,
Unprofiled RDEL BL recurring old costs,Finance & Benefits,E70,
Unprofiled RDEL BL Total,Finance & Benefits,F70,
Unprofiled Remainder RDEL Forecast - One off new costs - investment in change (GMPP - Remaining Spend),Finance & Benefits,C71,
Unprofiled RDEL Forecast recurring new costs,Finance & Benefits,D71,
Unprofiled RDEL Forecast recurring old costs,Finance & Benefits,E71,
Unprofiled RDEL Forecast Total,Finance & Benefits,F71,
Total RDEL BL one off new costs,Finance & Benefits,C72,
Total RDEL BL recurring new costs,Finance & Benefits,D72,
Total RDEL BL recurring old costs,Finance & Benefits,E72,
Total RDEL BL Total,Finance & Benefits,F72,
Total RDEL Forecast one off new costs,Finance & Benefits,C73,
Total RDEL Forecast recurring new costs,Finance & Benefits,D73,
Total RDEL Forecast recurring old costs,Finance & Benefits,E73,
Total RDEL Forecast Total,Finance & Benefits,F73,
Annual Steady State for RDEL recurring new costs,Finance & Benefits,D74,
Year RDEL spend stops,Finance & Benefits,C75,Years (Spend),
Pre 14-15 CDEL BL one off new costs,Finance & Benefits,C80,
Pre 14-15 CDEL BL recurring new costs,Finance & Benefits,D80,
Pre 14-15 CDEL BL recurring old costs,Finance & Benefits,E80,
Pre 14-15 CDEL BL Total,Finance & Benefits,F80,
Pre 14-15 CDEL Actual one off new costs,Finance & Benefits,C81,
Pre 14-15 CDEL Actual recurring new costs,Finance & Benefits,D81,
Pre 14-15 CDEL Actual recurring old costs,Finance & Benefits,E81,
Pre 14-15 CDEL Actual Total,Finance & Benefits,F81,
14-15 CDEL BL one off new costs,Finance & Benefits,C82,
14-15 CDEL BL recurring new costs,Finance & Benefits,D82,
14-15 CDEL BL recurring old costs,Finance & Benefits,E82,
14-15 CDEL BL Total,Finance & Benefits,F82,
14-15 CDEL Actual one off new costs,Finance & Benefits,C83,
14-15 CDEL Actual recurring new costs,Finance & Benefits,D83,
14-15 CDEL Actual recurring old costs,Finance & Benefits,E83,
14-15 CDEL Actual Total,Finance & Benefits,F83,
15-16 CDEL BL one off new costs,Finance & Benefits,C84,
15-16 CDEL BL recurring new costs,Finance & Benefits,D84,
15-16 CDEL BL recurring old costs,Finance & Benefits,E84,
15-16 CDEL BL Total,Finance & Benefits,F84,
15-16 CDEL Actual one off new costs,Finance & Benefits,C85,
15-16 CDEL Actual recurring new costs,Finance & Benefits,D85,
15-16 CDEL Actual recurring old costs,Finance & Benefits,E85,
15-16 CDEL Actual Total,Finance & Benefits,F85,
16-17 CDEL BL one off new costs,Finance & Benefits,C86,
16-17 CDEL BL recurring new costs,Finance & Benefits,D86,
16-17 CDEL BL recurring old costs,Finance & Benefits,E86,
16-17 CDEL BL Total,Finance & Benefits,F86,
16-17 CDEL Actual one off new costs,Finance & Benefits,C87,
16-17 CDEL Actual recurring new costs,Finance & Benefits,D87,
16-17 CDEL Actual recurring old costs,Finance & Benefits,E87,
16-17 CDEL Actual Total,Finance & Benefits,F87,
Pre 17-18 CDEL BL one off new costs,Finance & Benefits,C88,
Pre 17-18 CDEL BL recurring new costs,Finance & Benefits,D88,
Pre 17-18 CDEL BL recurring old costs,Finance & Benefits,E88,
Pre 17-18 CDEL BL Total,Finance & Benefits,F88,
Pre 17-18 CDEL Actual one off new costs,Finance & Benefits,C89,
Pre 17-18 CDEL Actual recurring new costs,Finance & Benefits,D89,
Pre 17-18 CDEL Actual recurring old costs,Finance & Benefits,E89,
Pre 17-18 CDEL Actual Total,Finance & Benefits,F89,
CDEL one off new cost spend in year on profile,Finance & Benefits,C90,
CDEL recurring new cost spend in year on profile,Finance & Benefits,D90,
CDEL recurring old cost spend in year on profile,Finance & Benefits,E90,
CDEL total spend in year on profile,Finance & Benefits,F90,
17-18 CDEL BL one off new costs,Finance & Benefits,C91,
17-18 CDEL BL recurring new costs,Finance & Benefits,D91,
17-18 CDEL BL recurring old costs,Finance & Benefits,E91,
17-18 CDEL BL Total,Finance & Benefits,F91,
17-18 CDEL Forecast one off new costs,Finance & Benefits,C92,
17-18 CDEL Forecast recurring new costs,Finance & Benefits,D92,
17-18 CDEL Forecast recurring old costs,Finance & Benefits,E92,
17-18 CDEL Forecast Total,Finance & Benefits,F92,
Pre 18-19 CDEL BL one off new costs,Finance & Benefits,C93,
Pre 18-19 CDEL BL recurring new costs,Finance & Benefits,D93,
Pre 18-19 CDEL BL recurring old costs,Finance & Benefits,E93,
Pre 18-19 CDEL BL Total,Finance & Benefits,F93,
Pre 18-19 CDEL Forecast one off new costs,Finance & Benefits,C94,
Pre 18-19 CDEL Forecast recurring new costs,Finance & Benefits,D94,
Pre 18-19 CDEL Forecast recurring old costs,Finance & Benefits,E94,
Pre 18-19 CDEL Forecast Total,Finance & Benefits,F94,
18-19 CDEL BL one off new costs,Finance & Benefits,C95,
18-19 CDEL BL recurring new costs,Finance & Benefits,D95,
18-19 CDEL BL recurring old costs,Finance & Benefits,E95,
18-19 CDEL BL Total,Finance & Benefits,F95,
18-19 CDEL Forecast one off new costs,Finance & Benefits,C96,
18-19 CDEL Forecast recurring new costs,Finance & Benefits,D96,
18-19 CDEL Forecast recurring old costs,Finance & Benefits,E96,
18-19 CDEL Forecast Total,Finance & Benefits,E96,
Pre 19-20 CDEL BL one off new costs,Finance & Benefits,C97,
Pre 19-20 CDEL BL recurring new costs,Finance & Benefits,D97,
Pre 19-20 CDEL BL recurring old costs,Finance & Benefits,E97,
Pre 19-20 CDEL BL Total,Finance & Benefits,F97,
Pre 19-20 CDEL Forecast one off new costs,Finance & Benefits,C98,
Pre 19-20 CDEL Forecast recurring new costs,Finance & Benefits,D98,
Pre 19-20 CDEL Forecast recurring old costs,Finance & Benefits,E98,
Pre 19-20 CDEL Forecast Total,Finance & Benefits,F98,
19-20 CDEL BL one off new costs,Finance & Benefits,C99,
19-20 CDEL BL recurring new costs,Finance & Benefits,D99,
19-20 CDEL BL recurring old costs,Finance & Benefits,E99,
19-20 CDEL BL Total,Finance & Benefits,F99,
19-20 CDEL Forecast one off new costs,Finance & Benefits,C100,
19-20 CDEL Forecast recurring new costs,Finance & Benefits,D100,
19-20 CDEL Forecast recurring old costs,Finance & Benefits,E100,
19-20 CDEL Forecast Total,Finance & Benefits,F100,
Pre 20-21 CDEL BL one off new costs,Finance & Benefits,C101,
Pre 20-21 CDEL BL recurring new costs,Finance & Benefits,D101,
Pre 20-21 CDEL BL recurring old costs,Finance & Benefits,E101,
Pre 20-21 CDEL BL Total,Finance & Benefits,F101,
Pre 20-21 CDEL Forecast one off new costs,Finance & Benefits,C102,
Pre 20-21 CDEL Forecast recurring new costs,Finance & Benefits,D102,
Pre 20-21 CDEL Forecast recurring old costs,Finance & Benefits,E102,
Pre 20-21 CDEL Forecast Total,Finance & Benefits,F102,
20-21 CDEL BL one off new costs,Finance & Benefits,C103,
20-21 CDEL BL recurring new costs,Finance & Benefits,D103,
20-21 CDEL BL recurring old costs,Finance & Benefits,E103,
20-21 CDEL BL Total,Finance & Benefits,F103,
20-21 CDEL Forecast one off new costs,Finance & Benefits,C104,
20-21 CDEL Forecast recurring new costs,Finance & Benefits,D104,
20-21 CDEL Forecast recurring old costs,Finance & Benefits,E104,
20-21 CDEL Forecast Total,Finance & Benefits,F104,
Pre 21-22 CDEL BL one off new costs,Finance & Benefits,C105,
Pre 21-22 CDEL BL recurring new costs,Finance & Benefits,D105,
Pre 21-22 CDEL BL recurring old costs,Finance & Benefits,E105,
Pre 21-22 CDEL BL Total,Finance & Benefits,F105,
Pre 21-22 CDEL Forecast one off new costs,Finance & Benefits,C106,
Pre 21-22 CDEL Forecast recurring new costs,Finance & Benefits,D106,
Pre 21-22 CDEL Forecast recurring old costs,Finance & Benefits,E106,
Pre 21-22 CDEL Forecast Total,Finance & Benefits,F106,
21-22 CDEL BL one off new costs,Finance & Benefits,C107,
21-22 CDEL BL recurring new costs,Finance & Benefits,D107,
21-22 CDEL BL recurring old costs,Finance & Benefits,E107,
21-22 CDEL BL Total,Finance & Benefits,F107,
21-22 CDEL Forecast one off new costs,Finance & Benefits,C108,
21-22 CDEL Forecast recurring new costs,Finance & Benefits,D108,
21-22 CDEL Forecast recurring old costs,Finance & Benefits,E108,
21-22 CDEL Forecast Total,Finance & Benefits,F108,
Pre 22-23 CDEL BL one off new costs,Finance & Benefits,C109,
Pre 22-23 CDEL BL recurring new costs,Finance & Benefits,D109,
Pre 22-23 CDEL BL recurring old costs,Finance & Benefits,E109,
Pre 22-23 CDEL BL Total,Finance & Benefits,F109,
Pre 22-23 CDEL Forecast one off new costs,Finance & Benefits,C110,
Pre 22-23 CDEL Forecast recurring new costs,Finance & Benefits,D110,
Pre 22-23 CDEL Forecast recurring old costs,Finance & Benefits,E110,
Pre 22-23 CDEL Forecast Total,Finance & Benefits,F110,
22-23 CDEL BL one off new costs,Finance & Benefits,C111,
22-23 CDEL BL recurring new costs,Finance & Benefits,D111,
22-23 CDEL BL recurring old costs,Finance & Benefits,E111,
22-23 CDEL BL Total,Finance & Benefits,F111,
22-23 CDEL Forecast one off new costs,Finance & Benefits,C112,
22-23 CDEL Forecast recurring new costs,Finance & Benefits,D112,
22-23 CDEL Forecast recurring old costs,Finance & Benefits,E112,
22-23 CDEL Forecast Total,Finance & Benefits,F112,
23-24 CDEL BL one off new costs,Finance & Benefits,C113,
23-24 CDEL BL recurring new costs,Finance & Benefits,D113,
23-24 CDEL BL recurring old costs,Finance & Benefits,E113,
23-24 CDEL BL Total,Finance & Benefits,F113,
23-24 CDEL Forecast one off new costs,Finance & Benefits,C114,
23-24 CDEL Forecast recurring new costs,Finance & Benefits,D114,
23-24 CDEL Forecast recurring old costs,Finance & Benefits,E114,
23-24 CDEL Forecast Total,Finance & Benefits,F114,
24-25 CDEL BL one off new costs,Finance & Benefits,C115,
24-25 CDEL BL recurring new costs,Finance & Benefits,D115,
24-25 CDEL BL recurring old costs,Finance & Benefits,E115,
24-25 CDEL BL Total,Finance & Benefits,F115,
24-25 CDEL Forecast one off new costs,Finance & Benefits,C116,
24-25 CDEL Forecast recurring new costs,Finance & Benefits,D116,
24-25 CDEL Forecast recurring old costs,Finance & Benefits,E116,
24-25 CDEL Forecast Total,Finance & Benefits,F116,
25-26 CDEL BL one off new costs,Finance & Benefits,C117,
25-26 CDEL BL recurring new costs,Finance & Benefits,D117,
25-26 CDEL BL recurring old costs,Finance & Benefits,E117,
25-26 CDEL BL Total,Finance & Benefits,F117,
25-26 CDEL Forecast one off new costs,Finance & Benefits,C118,
25-26 CDEL Forecast recurring new costs,Finance & Benefits,D118,
25-26 CDEL Forecast recurring old costs,Finance & Benefits,E118,
25-26 CDEL Forecast Total,Finance & Benefits,F118,
26-27 CDEL BL one off new costs,Finance & Benefits,C119,
26-27 CDEL BL recurring new costs,Finance & Benefits,D119,
26-27 CDEL BL recurring old costs,Finance & Benefits,E119,
26-27 CDEL BL Total,Finance & Benefits,F119,
26-27 CDEL Forecast one off new costs,Finance & Benefits,C120,
26-27 CDEL Forecast recurring new costs,Finance & Benefits,D120,
26-27 CDEL Forecast recurring old costs,Finance & Benefits,E120,
26-27 CDEL Forecast Total,Finance & Benefits,F120,
27-28 CDEL BL one off new costs,Finance & Benefits,C121,
27-28 CDEL BL recurring new costs,Finance & Benefits,D121,
27-28 CDEL BL recurring old costs,Finance & Benefits,E121,
27-28 CDEL BL Total,Finance & Benefits,F121,
27-28 CDEL Forecast one off new costs,Finance & Benefits,C122,
27-28 CDEL Forecast recurring new costs,Finance & Benefits,D122,
27-28 CDEL Forecast recurring old costs,Finance & Benefits,E122,
27-28 CDEL Forecast Total,Finance & Benefits,F122,
Unprofiled CDEL BL one off new costs,Finance & Benefits,C123,
Unprofiled CDEL BL recurring new costs,Finance & Benefits,D123,
Unprofiled CDEL BL recurring old costs,Finance & Benefits,E123,
Unprofiled CDEL BL Total,Finance & Benefits,F123,
Unprofiled CDEL Forecast one off new costs,Finance & Benefits,C124,
Unprofiled CDEL Forecast recurring new costs,Finance & Benefits,D124,
Unprofiled CDEL Forecast recurring old costs,Finance & Benefits,E124,
Unprofiled CDEL Forecast Total,Finance & Benefits,F124,
Total CDEL BL one off new costs,Finance & Benefits,C125,
Total CDEL BL recurring new costs,Finance & Benefits,D125,
Total CDEL BL recurring old costs,Finance & Benefits,E125,
Total CDEL BL Total,Finance & Benefits,F125,
Total CDEL Forecast one off new costs,Finance & Benefits,C126,
Total CDEL Forecast recurring new costs,Finance & Benefits,D126,
Total CDEL Forecast recurring old costs,Finance & Benefits,E126,
Total CDEL Forecast Total,Finance & Benefits,F126,
Annual Steady State for CDEL recurring new costs,Finance & Benefits,D127,
Year CDEL spend stops,Finance & Benefits,C128,Years (Spend),
Unprofiled Non-Gov BL,Finance & Benefits,H70,
Unprofiled Non-Gov Forecast,Finance & Benefits,H71,
Total BL Non-Gov,Finance & Benefits,H135,
Total Forecast Non-Gov,Finance & Benefits,I135,
Pre 14-15 BL Non-Gov,Finance & Benefits,H27,
Pre 14-15 Forecast Non-Gov BL,Finance & Benefits,H28,
14-15 BL Non-Gov,Finance & Benefits,H29,
14-15 Actual Non-Gov,Finance & Benefits,H30,
15-16 BL Non-Gov,Finance & Benefits,H31,
15-16 Actual Non-Gov,Finance & Benefits,H32,
16-17 BL Non-Gov,Finance & Benefits,H33,
16-17 Actual Non-Gov,Finance & Benefits,H34,
Pre 17-18 BL Non-Gov,Finance & Benefits,H35,
Pre 17-18 Actual Non-Gov,Finance & Benefits,H36,
2017/2018 Non-Gov (£m) Revenue & Capital Spend on profile?,Finance & Benefits,H37,
17-18 Forecast Non-Gov,Finance & Benefits,H38,
17-18 BL Non-Gov,Finance & Benefits,H39,
Pre 18-19 BL Non-Gov,Finance & Benefits,H40,
Pre 18-19 Forecast Non-Gov,Finance & Benefits,H41,
18-19 BL Non-Gov,Finance & Benefits,H42,
18-19 Forecast Non-Gov,Finance & Benefits,H43,
Pre 19-20 BL Non-Gov,Finance & Benefits,H44,
Pre 19-20 Forecast Non-Gov,Finance & Benefits,H45,
19-20 BL Non-Gov,Finance & Benefits,H46,
19-20 Forecast Non-Gov,Finance & Benefits,H47,
Pre 20-21 BL Non-Gov,Finance & Benefits,H48,
Pre 20-21 Forecast Non-Gov,Finance & Benefits,H49,
20-21 BL Non-Gov,Finance & Benefits,H50,
20-21 Forecast Non-Gov,Finance & Benefits,H51,
Pre 21-22 BL Non-Gov,Finance & Benefits,H52,
Pre 21-22 Forecast Non-Gov,Finance & Benefits,H53,
21-22 BL Non-Gov,Finance & Benefits,H54,
21-22 Forecast Non-Gov,Finance & Benefits,H55,
Pre 22-23 BL Non-Gov,Finance & Benefits,H56,
Pre 22-23 Forecast Non-Gov,Finance & Benefits,H57,
22-23 BL Non-Gov,Finance & Benefits,H58,
22-23 Forecast Non-Gov,Finance & Benefits,H59,
23-24 BL Non-Gov,Finance & Benefits,H60,
23-24 Forecast Non-Gov,Finance & Benefits,H61,
24-25 BL Non-Gov,Finance & Benefits,H62,
24-25 Forecast Non-Gov,Finance & Benefits,H63,
25-26 BL Non-Gov,Finance & Benefits,H64,
25-26 Forecast Non-Gov,Finance & Benefits,H65,
26-27 BL Non-Gov,Finance & Benefits,H66,
26-27 Forecast Non-Gov,Finance & Benefits,H67,
27-28 BL Non-Gov,Finance & Benefits,H68,
27-28 Forecast Non-Gov,Finance & Benefits,H69,
Unprofiled BL Non-Gov,Finance & Benefits,H70,
Unprofiled Forecast-Gov,Finance & Benefits,D71,
Pre 14-15 BL – Income both Revenue and Capital,Finance & Benefits,I27,
Pre 14-15 Actual – Income both Revenue and Capital,Finance & Benefits,I28,
14-15 BL - Income both Revenue and Capital,Finance & Benefits,I29,
14-15 Actual - Income both Revenue and Capital,Finance & Benefits,I30,
15-16 BL - Income both Revenue and Capital,Finance & Benefits,I31,
15-16 Actual - Income both Revenue and Capital,Finance & Benefits,I32,
16-17 BL - Income both Revenue and Capital,Finance & Benefits,I33,
16-17 Actual - Income both Revenue and Capital,Finance & Benefits,I34,
Pre 16-17 BL – Income both Revenue and Capital,Finance & Benefits,I35,
Pre 16-17 Actual – Income both Revenue and Capital,Finance & Benefits,I36,
2017/2018 Income Revenue & Capital Spend on Profile,Finance & Benefits,I37,
17-18 BL - Income both Revenue and Capital,Finance & Benefits,I38,
17-18 Forecast - Income both Revenue and Capital,Finance & Benefits,I39,
Pre 18-19 BL – Income both Revenue and Capital,Finance & Benefits,I40,
Pre 18-19 Actual – Income both Revenue and Capital,Finance & Benefits,I41,
18-19 BL - Income both Revenue and Capital,Finance & Benefits,I42,
18-19 Forecast - Income both Revenue and Capital,Finance & Benefits,I43,
Pre 19-20 BL – Income both Revenue and Capital,Finance & Benefits,I44,
Pre 19-20 Actual – Income both Revenue and Capital,Finance & Benefits,I45,
19-20 BL – Income both Revenue and Capital,Finance & Benefits,I46,
19-20 Forecast – Income both Revenue and Capital,Finance & Benefits,I47,
Pre 20-21 BL – Income both Revenue and Capital,Finance & Benefits,I48,
Pre 20-21 Actual – Income both Revenue and Capital,Finance & Benefits,I49,
20-21 BL - Income both Revenue and Capital,Finance & Benefits,I50,
20-21 Forecast – Income both Revenue and Capital,Finance & Benefits,I51,
Pre 21-22 BL – Income both Revenue and Capital,Finance & Benefits,I52,
Pre 21-22 Actual – Income both Revenue and Capital,Finance & Benefits,I53,
21-22 BL – Income both Revenue and Capital,Finance & Benefits,I54,
21-22 Forecast – Income both Revenue and Capital,Finance & Benefits,I55,
Pre 22-23 BL – Income both Revenue and Capital,Finance & Benefits,I56,
Pre 22-23 Actual – Income both Revenue and Capital,Finance & Benefits,I57,
22-23 BL – Income both Revenue and Capital,Finance & Benefits,I58,
22-23 Forecast – Income both Revenue and Capital,Finance & Benefits,I59,
23-24 BL – Income both Revenue and Capital,Finance & Benefits,I60,
23-24 Forecast – Income both Revenue and Capital,Finance & Benefits,I61,
24-25 BL – Income both Revenue and Capital,Finance & Benefits,I62,
24-25 Forecast – Income both Revenue and Capital,Finance & Benefits,I63,
25-26 BL – Income both Revenue and Capital,Finance & Benefits,I64,
25-26 Forecast – Income both Revenue and Capital,Finance & Benefits,I65,
26-27 BL – Income both Revenue and Capital,Finance & Benefits,I66,
26-27 Forecast – Income both Revenue and Capital,Finance & Benefits,I67,
27-28 BL – Income both Revenue and Capital,Finance & Benefits,I68,
27-28 Forecast – Income both Revenue and Capital,Finance & Benefits,I69,
Unprofiled BL Income,Finance & Benefits,I70,
Unprofiled Forecast Income,Finance & Benefits,I71,
Total Baseline - Income both Revenue and Capital,Finance & Benefits,I72,
Total Forecast - Income both Revenue and Capital,Finance & Benefits,I73,
Pre 14-15 BEN Baseline - Gov. Cashable,Finance & Benefits,C144,
Pre 14-15 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D144,
Pre 14-15 BEN Baseline - Economic (inc Economic (inc Private Partner),Finance & Benefits,E144,
Pre 14-15 BEN Baseline - Disbenefit Disbenefit UK Economic,Finance & Benefits,F144,
Pre 14-15 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G144,
Pre 14-15 BEN Actual - Gov. Cashable,Finance & Benefits,C155,
Pre 14-15 BEN Actual - Gov. Non-Cashable,Finance & Benefits,D155,
Pre 14-15 BEN Actual - Economic (inc Private Partner),Finance & Benefits,E155,
Pre 14-15 BEN Actual - Disbenefit UK Economic,Finance & Benefits,F155,
Pre 14-15 BEN Actual- Total Monetised Benefits,Finance & Benefits,G155,
14-15 BEN Baseline - Gov. Cashable,Finance & Benefits,C146,
14-15 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D146,
14-15 BEN Baseline - Economic (inc Economic (inc Private Partner),Finance & Benefits,E146,
14-15 BEN Baseline - Disbenefit Disbenefit UK Economic,Finance & Benefits,F146,
14-15 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G146,
14-15 BEN Actual - Gov. Cashable,Finance & Benefits,C147,
14-15 BEN Actual - Gov. Non-Cashable,Finance & Benefits,D147,
14-15 BEN Actual - Economic (inc Private Partner),Finance & Benefits,E147,
14-15 BEN Actual - Disbenefit UK Economic,Finance & Benefits,F147,
14-15 BEN Actual- Total Monetised Benefits,Finance & Benefits,G147,
15-16 BEN Baseline - Gov. Cashable,Finance & Benefits,C148,
15-16 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D148,
15-16 BEN Baseline - Economic (inc Economic (inc Private Partner),Finance & Benefits,E148,
15-16 BEN Baseline - Disbenefit Disbenefit UK Economic,Finance & Benefits,F148,
15-16 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G148,
15-16 BEN Actual - Gov. Cashable,Finance & Benefits,C149,
15-16 BEN Actual - Gov. Non-Cashable,Finance & Benefits,D149,
15-16 BEN Actual - Economic (inc Private Partner),Finance & Benefits,E149,
15-16 BEN Actual - Disbenefit UK Economic,Finance & Benefits,F149,
15-16 BEN Actual- Total Monetised Benefits,Finance & Benefits,G149,
16-17 BEN Baseline - Gov. Cashable,Finance & Benefits,C150,
16-17 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D150,
16-17 BEN Baseline - Economic (inc Economic (inc Private Partner),Finance & Benefits,E150,
16-17 BEN Baseline - Disbenefit Disbenefit UK Economic,Finance & Benefits,F150,
16-17 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G150,
16-17 BEN Actual - Gov. Cashable,Finance & Benefits,C151,
16-17 BEN Actual - Gov. Non-Cashable,Finance & Benefits,D151,
16-17 BEN Actual - Economic (inc Private Partner),Finance & Benefits,E151,
16-17 BEN Actual - Disbenefit UK Economic,Finance & Benefits,F151,
16-17 BEN Actual- Total Monetised Benefits,Finance & Benefits,G151,
Pre 17-18 BEN Baseline - Gov. Cashable,Finance & Benefits,C152,
Pre 17-18 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D152,
Pre 17-18 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E152,
Pre 17-18 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F152,
Pre 17-18 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G152,
Pre 17-18 BEN Forecast - Gov. Cashable,Finance & Benefits,C153,
Pre 17-18 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D153,
Pre 17-18 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E153,
Pre 17-18 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F153,
Pre 17-18 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G153,
17-18 BEN Baseline - Gov. Cashable,Finance & Benefits,C154,
17-18 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D154,
17-18 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E154,
17-18 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F154,
17-18 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G154,
17-18 BEN Forecast - Gov. Cashable,Finance & Benefits,C155,
17-18 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D155,
17-18 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E155,
17-18 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F155,
17-18 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G155,
Pre 18-19 BEN Baseline - Gov. Cashable,Finance & Benefits,C156,
Pre 18-19 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D156,
Pre 18-19 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E156,
Pre 18-19 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F156,
Pre 18-19 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G156,
Pre 18-19 BEN Forecast - Gov. Cashable,Finance & Benefits,C157,
Pre 18-19 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D157,
Pre 18-19 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E157,
Pre 18-19 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F157,
Pre 18-19 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G157,
18-19 BEN Baseline - Gov. Cashable,Finance & Benefits,C158,
18-19 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D158,
18-19 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E158,
18-19 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F158,
18-19 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G158,
18-19 BEN Forecast - Gov. Cashable,Finance & Benefits,C159,
18-19 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D159,
18-19 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E159,
18-19 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F159,
18-19 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G159,
Pre 19-20 BEN Baseline - Gov. Cashable,Finance & Benefits,C160,
Pre 19-20 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D160,
Pre 19-20 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E160,
Pre 19-20 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F160,
Pre 19-20 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G160,
Pre 19-20 BEN Forecast - Gov. Cashable,Finance & Benefits,C161,
Pre 19-20 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D161,
Pre 19-20 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E161,
Pre 19-20 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F161,
Pre 19-20 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G161,
19-20 BEN Baseline - Gov. Cashable,Finance & Benefits,C162,
19-20 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D162,
19-20 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E162,
19-20 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F162,
19-20 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G162,
19-20 BEN Forecast - Gov. Cashable,Finance & Benefits,C163,
19-20 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D163,
19-20 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E163,
19-20 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F163,
19-20 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G163,
Pre 20-21 BEN Baseline - Gov. Cashable,Finance & Benefits,C164,
Pre 20-21 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D164,
Pre 20-21 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E164,
Pre 20-21 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F164,
Pre 20-21 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G164,
Pre 20-21 BEN Forecast - Gov. Cashable,Finance & Benefits,C165,
Pre 20-21 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D165,
Pre 20-21 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E165,
Pre 20-21 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F165,
Pre 20-21 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G165,
20-21 BEN Baseline - Gov. Cashable,Finance & Benefits,C166,
20-21 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D166,
20-21 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E166,
20-21 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F166,
20-21 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G166,
20-21 BEN Forecast - Gov. Cashable,Finance & Benefits,C167,
20-21 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D167,
20-21 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E167,
20-21 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F167,
20-21 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G167,
Pre 21-22 BEN Baseline - Gov. Cashable,Finance & Benefits,C168,
Pre 21-22 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D168,
Pre 21-22 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E168,
Pre 21-22 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F168,
Pre 21-22 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G168,
Pre 21-22 BEN Forecast - Gov. Cashable,Finance & Benefits,C169,
Pre 21-22 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D169,
Pre 21-22 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E169,
Pre 21-22 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F169,
Pre 21-22 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G169,
21-22 BEN Baseline - Gov. Cashable,Finance & Benefits,C170,
21-22 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D170,
21-22 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E170,
21-22 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F170,
21-22 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G170,
21-22 BEN Forecast - Gov. Cashable,Finance & Benefits,C171,
21-22 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D171,
21-22 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E171,
21-22 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F171,
21-22 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G171,
Pre 22-23 BEN Baseline - Gov. Cashable,Finance & Benefits,C172,
Pre 22-23 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D172,
Pre 22-23 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E172,
Pre 22-23 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F172,
Pre 22-23 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G172,
Pre 22-23 BEN Forecast - Gov. Cashable,Finance & Benefits,C173,
Pre 22-23 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D173,
Pre 22-23 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E173,
Pre 22-23 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F173,
Pre 22-23 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G173,
22-23 BEN Baseline - Gov. Cashable,Finance & Benefits,C174,
22-23 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D174,
22-23 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E174,
22-23 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F174,
22-23 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G174,
22-23 BEN Forecast - Gov. Cashable,Finance & Benefits,C175,
22-23 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D175,
22-23 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E175,
22-23 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F175,
22-23 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G175,
23-24 BEN Baseline - Gov. Cashable,Finance & Benefits,C176,
23-24 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D176,
23-24 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E176,
23-24 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F176,
23-24 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G176,
23-24 BEN Forecast - Gov. Cashable,Finance & Benefits,C177,
23-24 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D177,
23-24 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E177,
23-24 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F177,
23-24 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G177,
24-25 BEN Baseline - Gov. Cashable,Finance & Benefits,C178,
24-25 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D178,
24-25 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E178,
24-25 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F178,
24-25 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G178,
24-25 BEN Forecast - Gov. Cashable,Finance & Benefits,C179,
24-25 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D179,
24-25 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E179,
24-25 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F179,
24-25 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G179,
25-26 BEN Baseline - Gov. Cashable,Finance & Benefits,C180,
25-26 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D180,
25-26 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E180,
25-26 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F180,
25-26 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G180,
25-26 BEN Forecast - Gov. Cashable,Finance & Benefits,C181,
25-26 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D181,
25-26 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E181,
25-26 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F181,
25-26 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G181,
26-27 BEN Baseline - Gov. Cashable,Finance & Benefits,C182,
26-27 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D182,
26-27 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E182,
26-27 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F182,
26-27 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G182,
26-27 BEN Forecast - Gov. Cashable,Finance & Benefits,C183,
26-27 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D183,
26-27 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E183,
26-27 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F183,
26-27 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G183,
27-28 BEN Baseline - Gov. Cashable,Finance & Benefits,C184,
27-28 BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D184,
27-28 BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E184,
27-28 BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F184,
27-28 BEN Baseline - Total Monetised Benefits,Finance & Benefits,G184,
27-28 BEN Forecast - Gov. Cashable,Finance & Benefits,C185,
27-28 BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D185,
27-28 BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E185,
27-28 BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F185,
27-28 BEN Forecast - Total Monetised Benefits,Finance & Benefits,G185,
Unprofiled Remainder BEN Baseline - Gov. Cashable,Finance & Benefits,C186,
Unprofiled Remainder BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D186,
Unprofiled Remainder BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E186,
Unprofiled Remainder BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F186,
Unprofiled Remainder BEN Baseline - Total Monetised Benefits,Finance & Benefits,G186,
Unprofiled Remainder BEN Forecast - Gov. Cashable,Finance & Benefits,C187,
Unprofiled Remainder BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D187,
Unprofiled Remainder BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E187,
Unprofiled Remainder BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F187,
Unprofiled Remainder BEN Forecast - Total Monetised Benefits,Finance & Benefits,G187,
Total BEN Baseline - Gov. Cashable,Finance & Benefits,C188,
Total BEN Baseline - Gov. Non-Cashable,Finance & Benefits,D188,
Total BEN Baseline - Economic (inc Private Partner),Finance & Benefits,E188,
Total BEN Baseline - Disbenefit UK Economic,Finance & Benefits,F188,
Total BEN Baseline - Total Monetised Benefits,Finance & Benefits,G188,
Total BEN Forecast - Gov. Cashable,Finance & Benefits,C189,
Total BEN Forecast - Gov. Non-Cashable,Finance & Benefits,D189,
Total BEN Forecast - Economic (inc Private Partner),Finance & Benefits,E189,
Total BEN Forecast - Disbenefit UK Economic,Finance & Benefits,F189,
Total BEN Forecast - Total Monetised Benefits,Finance & Benefits,G189,
Year BEN spend stops,Finance & Benefits,C191,Years (Benefits),
Benefits Narrative,Finance & Benefits,B197,
Benefits Map,Finance & Benefits,F139,Yes/No,
Benefits Analysed,Finance & Benefits,F140,Yes/No,
Benefits Realisation Plan,Finance & Benefits,F141,Yes/No,
Initial Benefits Cost Ratio (BCR),Finance & Benefits,F191,
Adjusted Benefits Cost Ratio (BCR),Finance & Benefits,F192,
VfM Category,Finance & Benefits,B192,VFM,
Present Value Cost (PVC),Finance & Benefits,I191,
Present Value Benefit (PVB),Finance & Benefits,I192,
Assurance MM1,Assurance Planning,A8,
Assurance MM1 Original Baseline,Assurance Planning,B8,
Assurance MM1 Latest Approved Baseline,Assurance Planning,C8,
Assurance MM1 Forecast - Actual,Assurance Planning,D8,
Assurance MM1 DCA,Assurance Planning,E8,RAG,
Assurance MM1 Notes,Assurance Planning,F8,
Assurance MM2,Assurance Planning,A9,
Assurance MM2 Original Baseline,Assurance Planning,B9,
Assurance MM2 Latest Approved Baseline,Assurance Planning,C9,
Assurance MM2 Forecast - Actual,Assurance Planning,D9,
Assurance MM2 DCA,Assurance Planning,E9,RAG,
Assurance MM2 Notes,Assurance Planning,F9,
Assurance MM3,Assurance Planning,A10,
Assurance MM3 Original Baseline,Assurance Planning,B10,
Assurance MM3 Latest Approved Baseline,Assurance Planning,C10,
Assurance MM3 Forecast - Actual,Assurance Planning,D10,
Assurance MM3 DCA,Assurance Planning,E10,RAG,
Assurance MM3 Notes,Assurance Planning,F10,
Assurance MM4,Assurance Planning,A11,
Assurance MM4 Original Baseline,Assurance Planning,B11,
Assurance MM4 Latest Approved Baseline,Assurance Planning,C11,
Assurance MM4 Forecast - Actual,Assurance Planning,D11,
Assurance MM4 DCA,Assurance Planning,E11,RAG,
Assurance MM4 Notes,Assurance Planning,F11,
Assurance MM5,Assurance Planning,A12,
Assurance MM5 Original Baseline,Assurance Planning,B12,
Assurance MM5 Latest Approved Baseline,Assurance Planning,C12,
Assurance MM5 Forecast - Actual,Assurance Planning,D12,
Assurance MM5 DCA,Assurance Planning,E12,RAG,
Assurance MM5 Notes,Assurance Planning,F12,
Assurance MM6,Assurance Planning,A13,
Assurance MM6 Original Baseline,Assurance Planning,B13,
Assurance MM6 Latest Approved Baseline,Assurance Planning,C13,
Assurance MM6 Forecast - Actual,Assurance Planning,D13,
Assurance MM6 DCA,Assurance Planning,E13,RAG,
Assurance MM6 Notes,Assurance Planning,F13,
Assurance MM7,Assurance Planning,A14,
Assurance MM7 Original Baseline,Assurance Planning,B14,
Assurance MM7 Latest Approved Baseline,Assurance Planning,C14,
Assurance MM7 Forecast - Actual,Assurance Planning,D14,
Assurance MM7 DCA,Assurance Planning,E14,RAG,
Assurance MM7 Notes,Assurance Planning,F14,
Assurance MM8,Assurance Planning,A15,
Assurance MM8 Original Baseline,Assurance Planning,B15,
Assurance MM8 Latest Approved Baseline,Assurance Planning,C15,
Assurance MM8 Forecast - Actual,Assurance Planning,D15,
Assurance MM8 DCA,Assurance Planning,E15,RAG,
Assurance MM8 Notes,Assurance Planning,F15,
Assurance MM9,Assurance Planning,A16,
Assurance MM9 Original Baseline,Assurance Planning,B16,
Assurance MM9 Latest Approved Baseline,Assurance Planning,C16,
Assurance MM9 Forecast - Actual,Assurance Planning,D16,
Assurance MM9 DCA,Assurance Planning,E16,RAG,
Assurance MM9 Notes,Assurance Planning,F16,
Assurance MM10,Assurance Planning,A17,
Assurance MM10 Original Baseline,Assurance Planning,B17,
Assurance MM10 Latest Approved Baseline,Assurance Planning,C17,
Assurance MM10 Forecast - Actual,Assurance Planning,D17,
Assurance MM10 DCA,Assurance Planning,E17,RAG,
Assurance MM10 Notes,Assurance Planning,F17,
Assurance MM11,Assurance Planning,A18,
Assurance MM11 Original Baseline,Assurance Planning,B18,
Assurance MM11 Latest Approved Baseline,Assurance Planning,C18,
Assurance MM11 Forecast - Actual,Assurance Planning,D18,
Assurance MM11 DCA,Assurance Planning,E18,RAG,
Assurance MM11 Notes,Assurance Planning,F18,
Assurance MM12,Assurance Planning,A19,
Assurance MM12 Original Baseline,Assurance Planning,B19,
Assurance MM12 Latest Approved Baseline,Assurance Planning,C19,
Assurance MM12 Forecast - Actual,Assurance Planning,D19,
Assurance MM12 DCA,Assurance Planning,E19,RAG,
Assurance MM12 Notes,Assurance Planning,F19,
Assurance MM13,Assurance Planning,A20,
Assurance MM13 Original Baseline,Assurance Planning,B20,
Assurance MM13 Latest Approved Baseline,Assurance Planning,C20,
Assurance MM13 Forecast - Actual,Assurance Planning,D20,
Assurance MM13 DCA,Assurance Planning,E20,RAG,
Assurance MM13 Notes,Assurance Planning,F20,
Assurance MM14,Assurance Planning,A21,
Assurance MM14 Original Baseline,Assurance Planning,B21,
Assurance MM14 Latest Approved Baseline,Assurance Planning,C21,
Assurance MM14 Forecast - Actual,Assurance Planning,D21,
Assurance MM14 DCA,Assurance Planning,E21,RAG,
Assurance MM14 Notes,Assurance Planning,F21,
Assurance MM15,Assurance Planning,A22,
Assurance MM15 Original Baseline,Assurance Planning,B22,
Assurance MM15 Latest Approved Baseline,Assurance Planning,C22,
Assurance MM15 Forecast - Actual,Assurance Planning,D22,
Assurance MM15 DCA,Assurance Planning,E22,RAG,
Assurance MM15 Notes,Assurance Planning,F22,
Assurance MM16,Assurance Planning,A23,
Assurance MM16 Original Baseline,Assurance Planning,B23,
Assurance MM16 Latest Approved Baseline,Assurance Planning,C23,
Assurance MM16 Forecast - Actual,Assurance Planning,D23,
Assurance MM16 DCA,Assurance Planning,E23,RAG,
Assurance MM16 Notes,Assurance Planning,F23,
Assurance MM17,Assurance Planning,A24,
Assurance MM17 Original Baseline,Assurance Planning,B24,
Assurance MM17 Latest Approved Baseline,Assurance Planning,C24,
Assurance MM17 Forecast - Actual,Assurance Planning,D24,
Assurance MM17 DCA,Assurance Planning,E24,RAG,
Assurance MM17 Notes,Assurance Planning,F24,
Assurance MM18,Assurance Planning,A25,
Assurance MM18 Original Baseline,Assurance Planning,B25,
Assurance MM18 Latest Approved Baseline,Assurance Planning,C25,
Assurance MM18 Forecast - Actual,Assurance Planning,D25,
Assurance MM18 DCA,Assurance Planning,E25,RAG,
Assurance MM18 Notes,Assurance Planning,F25,
IAAP created,Assurance Planning,C4,
IAAP date revised,Assurance Planning,C5,
IAPP version,Assurance Planning,E4,
SRO assurance confidence RAG internal,Assurance Planning,C31,RAG,
SRO assurance confidence RAG external,Assurance Planning,C32,RAG,
SRO assurance confidence commentary,Assurance Planning,D30,
SRO assurance scope RAG internal,Assurance Planning,C36,RAG,
SRO assurance scope RAG external,Assurance Planning,C37,RAG,
SRO assurance scope RAG commentary,Assurance Planning,D35,
SRO Benefits RAG,Finance & Benefits,C141,RAG 2,
Total Number of public sector employees working on the project,Resource,C37,
Total Number of external contractors working on the project,Resource,E37,
Total Number or vacancies on the project,Resource,G37,
Resources commentary,Resource,C19,
Total number of employees funded to work on project,Resource,I17,
Resources commentary,Resource,C19,
Overall Resource DCA - Now,Resource,I38,Capability RAG,
Overall Resource DCA - Future,Resource,J38,Capability RAG,
Digital - Now,Resource,I25,Capability RAG,
Digital - Future,Resource,J25,Capability RAG,
Information Technology - Now,Resource,I26,Capability RAG,
Information Technology - Future,Resource,J26,Capability RAG,
Legal Commercial Contract Management - Now,Resource,I27,Capability RAG,
Legal Commercial Contract Management - Future,Resource,J27,Capability RAG,
Project Delivery - Now,Resource,I28,Capability RAG,
Project Delivery - Future,Resource,J28,Capability RAG,
Change Implementation - Now,Resource,I29,Capability RAG,
Change Implementation - Future,Resource,J29,Capability RAG,
Technical - Now,Resource,I30,Capability RAG,
Technical - Future,Resource,J30,Capability RAG,
Industry Knowledge - Now,Resource,I31,Capability RAG,
Industry Knowledge - Future,Resource,J31,Capability RAG,
Finance - Now,Resource,I32,Capability RAG,
Finance - Future,Resource,J32,Capability RAG,
Analysis Now,Resource,I33,Capability RAG,
Analysis - future,Resource,J33,Capability RAG,
Communications & Stakeholder Engagement - Now,Resource,I34,Capability RAG,
Communications & Stakeholder Engagement - Future,Resource,J34,Capability RAG,
Other Capability 3,Resource,A35,
Other Capability 3 - Now,Resource,I35,Capability RAG,
Other Capability 3 - Future,Resource,J35,Capability RAG,
Other Capability 4,Resource,A36,
Other Capability 4 - Now,Resource,I36,Capability RAG,
Other Capability 4 - Future,Resource,J36,Capability RAG,
Cap Commentary,Resource,C39,
SCS PB3 No public sector,Resource,C6,
SCS PB3 No externals,Resource,E6,
SCS PB3 No vacancies,Resource,G6,
SCS PB3 Total,Resource,I6,
SCS PB2 No public sector,Resource,C7,
SCS PB2 No externals,Resource,E7,
SCS PB2 No vacancies,Resource,G7,
SCS PB2 Total,Resource,I7,
SCS PB1 No public sector,Resource,C8,
SCS PB1 No externals,Resource,E8,
SCS PB1 No vacancies,Resource,G8,
SCS PB1 Total,Resource,I8,
G6 No public sector,Resource,C9,
G6 No externals,Resource,E9,
G6 No vacancies,Resource,G9,
G6 Total,Resource,I9,
G7 No public sector,Resource,C10,
G7 No externals,Resource,E10,
G7 No vacancies,Resource,G10,
G7 Total,Resource,I10,
FS No public sector,Resource,C11,
FS No externals,Resource,E11,
FS No vacancies,Resource,G11,
FS Total,Resource,I11,
SEO No public sector,Resource,C12,
SEO No externals,Resource,E12,
SEO No vacancies,Resource,G12,
SEO Total,Resource,I12,
HEO No public sector,Resource,C13,
HEO No externals,Resource,E13,
HEO No vacancies,Resource,G13,
HEO Total,Resource,I13,
EO No public sector,Resource,C14,
EO No externals,Resource,E14,
EO No vacancies,Resource,G14,
EO Total,Resource,I14,
AO No public sector,Resource,C15,
AO No externals,Resource,E15,
AO No vacancies,Resource,G15,
AO Total,Resource,I15,
AA No public sector,Resource,C16,
AA No externals,Resource,E16,
AA No vacancies,Resource,G16,
AA Total,Resource,I16,
Digital No public sector,Resource,C25,
Digital No externals,Resource,E25,
Digital No vacancies,Resource,G25,
IT No public sector,Resource,C26,
IT No externals,Resource,E26,
IT No vacancies,Resource,G26,
Legal Commercial No public sector,Resource,C27,
Legal Commercial No externals,Resource,E27,
Legal Commercial No vacancies,Resource,G27,
PD No public sector,Resource,C28,
PD No externals,Resource,E28,
PD No vacancies,Resource,G28,
Change Implementation No public sector,Resource,C29,
Change Implementation No externals,Resource,E29,
Change Implementation No vacancies,Resource,G29,
Technical No public sector,Resource,C30,
Technical No externals,Resource,E30,
Technical No vacancies,Resource,G30,
Industry Knowledge No public sector,Resource,C31,
Industry Knowledge No externals,Resource,E31,
Industry Knowledge No vacancies,Resource,G31,
Finance No public sector,Resource,C32,
Finance No externals,Resource,E32,
Finance No vacancies,Resource,G32,
Analysis No public sector,Resource,C33,
Analysis externals,Resource,E33,
Analysis No vacancies,Resource,G33,
Communications and Stakeholder No public sector,Resource,C34,
Communications and Stakeholder No externals,Resource,E34,
Communications and Stakeholder No vacancies,Resource,G34,
Other 3 No public sector,Resource,C35,
Other 3 No externals,Resource,E35,
Other 3 No vacancies,Resource,G35,
Other 4 No public sector,Resource,C36,
Other 4 No externals,Resource,E36,
Other 4 No vacancies,Resource,G36,
Total No public sector employees working on the project,Resource,C17,
Total No external contractors working on the project,Resource,E17,
Total No employees funded to work on project,Resource,I17,
GMPP - IPA ID Number,GMPP,B1,
GMPP - IPA ID Number 2,GMPP,B2,
GMPP - Dept,GMPP,B3,
GMPP - Main reason for joining GMPP,GMPP,B4,
GMPP - IPA DCA,GMPP,B5,
GMPP - IPA DCA Commentary,GMPP,B6,
GMPP - SRO ID,GMPP,B7,
GMPP - PD ID,GMPP,B8,
GMPP Annual Report Category,GMPP,B9,
GMPP Quarter ID,GMPP,B10,
"""

datamap_data = """
Project/Programme Name,Summary,B5,
SRO Sign-Off,Summary,B49,
Reporting period (GMPP - Snapshot Date),Summary,G3,
Quarter Joined,Summary,I3,
GMPP (GMPP - formally joined GMPP),Summary,G5,
IUK top 40,Summary,G6,
Top 37,Summary,I5,
DfT Business Plan,Summary,I6,
DFT ID Number,Summary,B6,
MPA ID Number,Summary,C6,
Working Contact Name,Summary,H8,
Working Contact Telephone,Summary,H9,
SRO Tenure Start Date,Summary,C15,
SRO Tenure End Date,Summary,C17,
Working Contact Email,Summary,H10,
DfT Group,Summary,B8,DfT Group,
DfT Division,Summary,B9,DfT Division,
Agency or delivery partner (GMPP - Delivery Organisation primary),Summary,B10,Agency,
Strategic Alignment/Government Policy (GMPP - Key drivers),Summary,B26,
Project stage,Approval & Project milestones,B5,Project stage,
Project stage if Other,Approval & Project milestones,D5,
Last time at BICC,Approval & Project milestones,B4,
Next at BICC,Approval & Project milestones,D4,
Approval MM1,Approval & Project milestones,A9,
Approval MM1 Original Baseline,Approval & Project milestones,B9,
Approval MM1 Latest Approved Baseline,Approval & Project milestones,C9,
Approval MM1 Forecast / Actual,Approval & Project milestones,D9,
Approval MM1 Milestone Type,Approval & Project milestones,E9,Milestone Types,
Approval MM1 Notes,Approval & Project milestones,F9,
Approval MM2,Approval & Project milestones,A10,
Approval MM2 Original Baseline,Approval & Project milestones,B10,
Approval MM2 Latest Approved Baseline,Approval & Project milestones,C10,
Approval MM2 Forecast / Actual,Approval & Project milestones,D10,
Approval MM2 Milestone Type,Approval & Project milestones,E10,
Approval MM2 Notes,Approval & Project milestones,F10,
Approval MM3,Approval & Project milestones,A11,
Approval MM3 Original Baseline,Approval & Project milestones,B11,
Approval MM3 Latest Approved Baseline,Approval & Project milestones,C11,
Approval MM3 Forecast / Actual,Approval & Project milestones,D11,
Approval MM3 Milestone Type,Approval & Project milestones,E11,Milestone Types,
Approval MM3 Notes,Approval & Project milestones,F11,
Significant Steel Requirement,Finance & Benefits,D15,Yes/No,
SRO Finance confidence,Finance & Benefits,C6,RAG 2,
BICC approval point,Finance & Benefits,E9,Business Cases,
Latest Treasury Approval Point (TAP) or equivalent,Finance & Benefits,E10,Business Cases,
Business Case used to source figures (GMPP TAP used to source figures),Finance & Benefits,C9,Business Cases,
Date of TAP used to source figures,Finance & Benefits,E11,
Name of source in not Business Case (GMPP -If not TAP please specify equivalent document used),Finance & Benefits,C10,
If not TAP please specify date of equivalent document,Finance & Benefits,C11,
Version Number Of Document used to Source Figures (GMPP - TAP version Number),Finance & Benefits,C12,
Date document approved by SRO,Finance & Benefits,C13,
Real or Nominal - Baseline,Finance & Benefits,C18,Finance figures format,
Real or Nominal - Actual/Forecast,Finance & Benefits,E18,Finance figures format,
Index Year,Finance & Benefits,B19,Index Years,
Deflator,Finance & Benefits,B20,Finance type,
Source of Finance,Finance & Benefits,B21,Finance type,
Other Finance type Description,Finance & Benefits,D21,
NPV for all projects and NPV for programmes if available,Finance & Benefits,B22,
Project cost to closure,Finance & Benefits,B23,
RDEL Total Budget/BL,Finance & Benefits,C72,
CDEL Total Budget/BL,Finance & Benefits,C125,
Non-Gov Total Budget/BL,Finance & Benefits,C135,
Total Budget/BL,Finance & Benefits,C136,
RDEL Total Forecast,Finance & Benefits,D133,
CDEL Total Forecast,Finance & Benefits,D134,
Non-Gov Total Forecast,Finance & Benefits,D135,
Total Forecast,Finance & Benefits,D136,
RDEL Total Variance,Finance & Benefits,E133,
CDEL Total Variance,Finance & Benefits,E134,
Assurance MM1,Assurance Planning,A8,
Assurance MM1 Original Baseline,Assurance Planning,B8,
Assurance MM1 Latest Approved Baseline,Assurance Planning,C8,
Assurance MM1 Forecast - Actual,Assurance Planning,D8,
Assurance MM1 DCA,Assurance Planning,E8,RAG,
Assurance MM1 Notes,Assurance Planning,F8,
Assurance MM2,Assurance Planning,A9,
Assurance MM2 Original Baseline,Assurance Planning,B9,
Assurance MM2 Latest Approved Baseline,Assurance Planning,C9,
Assurance MM2 Forecast - Actual,Assurance Planning,D9,
Assurance MM2 DCA,Assurance Planning,E9,RAG,
Assurance MM2 Notes,Assurance Planning,F9,
Total Number of public sector employees working on the project,Resource,C37,
Total Number of external contractors working on the project,Resource,E37,
Total Number or vacancies on the project,Resource,G37,
Resources commentary,Resource,C19,
Total number of employees funded to work on project,Resource,I17,
Resources commentary,Resource,C19,
Overall Resource DCA - Now,Resource,I38,Capability RAG,
Overall Resource DCA - Future,Resource,J38,Capability RAG,
Digital - Now,Resource,I25,Capability RAG,
Digital - Future,Resource,J25,Capability RAG,
Information Technology - Now,Resource,I26,Capability RAG,
Information Technology - Future,Resource,J26,Capability RAG,
Legal Commercial Contract Management - Now,Resource,I27,Capability RAG,
Legal Commercial Contract Management - Future,Resource,J27,Capability RAG,
Project Delivery - Now,Resource,I28,Capability RAG,
Project Delivery - Future,Resource,J28,Capability RAG,
Change Implementation - Now,Resource,I29,Capability RAG,
Change Implementation - Future,Resource,J29,Capability RAG,
Technical - Now,Resource,I30,Capability RAG,
Technical - Future,Resource,J30,Capability RAG,
Industry Knowledge - Now,Resource,I31,Capability RAG,
Industry Knowledge - Future,Resource,J31,Capability RAG,
Finance - Now,Resource,I32,Capability RAG,
Finance - Future,Resource,J32,Capability RAG,
Analysis Now,Resource,I33,Capability RAG,
Analysis - future,Resource,J33,Capability RAG,
"""


@pytest.fixture(scope='session')
def blank_template():
    gen_template(BICC_TEMPLATE_FOR_TESTS, SOURCE_DIR)
    output_file = '/'.join([SOURCE_DIR, 'gen_bicc_template.xlsm'])
#   yield output_file
    return output_file
#   os.remove(output_file)


@pytest.fixture(scope='session')
def datamap():
    name = 'datamap.csv'
    s = io.StringIO()
    s.write(datamap_header)
    s.write(real_datamap_data)
    s.seek(0)
    s_string = s.readlines()
#   del s_string[0]
    with open('/'.join([SOURCE_DIR, name]), 'w') as csv_file:
        for x in s_string:
            csv_file.write(x)
    return '/'.join([SOURCE_DIR, name])


@pytest.fixture
def populated_template():
    gen_template(BICC_TEMPLATE_FOR_TESTS, SOURCE_DIR)
    datamap()
    dm = "/".join([SOURCE_DIR, 'datamap.csv'])
    wb = load_workbook("/".join([SOURCE_DIR, 'gen_bicc_template.xlsm']), keep_vba=True)
    output_file = "/".join([RETURNS_DIR, 'populated_test_template.xlsm'])
    for fl in range(10):
        with open(dm, 'r', newline='') as f:
            reader = csv.DictReader(f)
            for line in reader:
                if line['cell_key'].startswith('Date'):  # we want to test date strings
                    wb[line['template_sheet']][line['cell_reference']].value = "20/06/2017"
                elif line['cell_key'].startswith('SRO Tenure'):  # we want to test date strings
                    wb[line['template_sheet']][line['cell_reference']].value = "10/08/2017"
                else:
                    wb[line['template_sheet']][line['cell_reference']].value = " ".join([line['cell_key'].upper(), str(fl)])
            output_file = "/".join([RETURNS_DIR, 'populated_test_template{}.xlsm'
                                    .format(fl)])
            wb.save(output_file)
    # we save 10 of them but only return the first for testing
    yield output_file
    fs = [f for f in os.listdir(RETURNS_DIR)]
    for f in fs:
        os.remove(os.path.join(RETURNS_DIR, f))



@pytest.fixture
def populated_template_comparison():
    gen_template(BICC_TEMPLATE_FOR_TESTS, SOURCE_DIR)
    datamap()
    dm = "/".join([SOURCE_DIR, 'datamap.csv'])
    wb = load_workbook("/".join([SOURCE_DIR, 'gen_bicc_template.xlsm']), keep_vba=True)
    output_file = "/".join([RETURNS_DIR, 'populated_test_template.xlsm'])
    for fl in range(3):
        with open(dm, 'r', newline='') as f:
            reader = csv.DictReader(f)
            for line in reader:
                if line['cell_key'].startswith('Date'):  # we want to test date strings
                    wb[line['template_sheet']][line['cell_reference']].value = "20/06/2017"
                elif line['cell_key'].startswith('SRO Tenure'):  # we want to test date strings
                    wb[line['template_sheet']][line['cell_reference']].value = "10/08/2017"
                else:
                    wb[line['template_sheet']][line['cell_reference']].value = " ".join([line['cell_key'].upper(), str(fl)])
            output_file = "/".join([RETURNS_DIR, 'populated_test_template{}.xlsm'
                                    .format(fl)])
            wb.save(output_file)
    # we save 3 of them but only return the first for testing
    yield output_file
    fs = [f for f in os.listdir(RETURNS_DIR)]
    for f in fs:
        os.remove(os.path.join(RETURNS_DIR, f))


def split_datamap_line(line: tuple):
    for item in line:
        yield item


@pytest.fixture(scope='session')
def master():
    """
    This is master file created for the purpose of using a base for bcompiler -a, which
    populates all the returns. It simply takes the field name from the datamap and
    puts it in upper case and appends a digit (1, 2 or 3 because we're only simulating
    a master with 3 projects here.
    :return: output_file
    """
    # regexes
    r'(Assurance MM1 .+$|Approval MM1 .+$)'
    milestones_regex1 = re.compile(r'(Assurance MM1 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM1 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex2 = re.compile(r'(Assurance MM2 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM2 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex4 = re.compile(r'(Assurance MM3 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM3 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex5 = re.compile(r'(Assurance MM4 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM4 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex6 = re.compile(r'(Assurance MM5 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM5 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex3 = re.compile(r'(Assurance MM\d+ (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM\d+ (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')

    wb = Workbook()
    output_file = "/".join([OUTPUT_DIR, 'master.xlsx'])
    ws = wb.active
    ws.title = "Master for Testing"
    for item in enumerate(real_datamap_data.split('\n')):
        if not item[0] == 0 and not item[1] == "":
            g = split_datamap_line(item)
            next(g)
            ix = next(g).split(',')[0]
            ws[f"A{str(item[0])}"] = ix
            if item[1].startswith('Date'):  # testing for date objects
                ws[f"B{str(item[0])}"] = date(2017, 6, 20)
                ws[f"C{str(item[0])}"] = date(2017, 6, 20)
                ws[f"D{str(item[0])}"] = date(2017, 6, 20)
            elif item[1].startswith('SRO Tenure'):  # testing for date objects
                ws[f"B{str(item[0])}"] = date(2017, 8, 10)
                ws[f"C{str(item[0])}"] = date(2017, 8, 10)
                ws[f"D{str(item[0])}"] = date(2017, 8, 10)
            elif item[1].startswith('Total Forecast'):
                ws[f"B{str(item[0])}"] = 32.3333
                ws[f"C{str(item[0])}"] = 35.2322
                ws[f"D{str(item[0])}"] = 23.2
            elif item[1].startswith('BICC approval point'):
                ws[f"B{str(item[0])}"] = "Strategic Outline Case"
                ws[f"C{str(item[0])}"] = "Outline Business Case"
                ws[f"D{str(item[0])}"] = "Full Business Case"
            elif item[1].startswith('Project MM20 Forecast - Actual'):
                ws[f"B{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"C{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"D{str(item[0])}"] = datetime(2016, 1, 1)
            elif item[1].startswith('Departmental DCA'):
                ws[f"B{str(item[0])}"] = "Amber/Red"
                ws[f"C{str(item[0])}"] = "Red"
                ws[f"D{str(item[0])}"] = "Green"
            elif item[1].startswith('SRO Finance confidence'):
                ws[f"B{str(item[0])}"] = "Amber/Red"
                ws[f"C{str(item[0])}"] = "Amber/Red"
                ws[f"D{str(item[0])}"] = "Amber"
            elif item[1].startswith('SRO Benefits RAG'):
                ws[f"B{str(item[0])}"] = "Red"
                ws[f"C{str(item[0])}"] = "Green"
                ws[f"D{str(item[0])}"] = "Amber/Green"
            elif item[1].startswith('GMPP - IPA DCA'):
                ws[f"B{str(item[0])}"] = "Amber"
                ws[f"C{str(item[0])}"] = "Amber"
                ws[f"D{str(item[0])}"] = "Amber/Green"


            # Here we are starting a block of dates. We need these to be able to test
            # the default swimlane_milstones analyser
            # we're giving these ones some variety so they can be tested as the default
            # swimlane_milestones analyser
            elif milestones_regex1.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2015, 1, 1)
                ws[f"C{str(item[0])}"] = date(2015, 1, 1)
                ws[f"D{str(item[0])}"] = date(2015, 1, 1)
            elif milestones_regex2.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2019, 1, 1)
                ws[f"C{str(item[0])}"] = date(2019, 1, 1)
                ws[f"D{str(item[0])}"] = date(2019, 1, 1)
            elif milestones_regex3.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2020, 1, 1)
                ws[f"C{str(item[0])}"] = date(2020, 1, 1)
                ws[f"D{str(item[0])}"] = date(2020, 1, 1)
            elif milestones_regex4.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2018, 1, 1)
                ws[f"C{str(item[0])}"] = date(2018, 1, 1)
                ws[f"D{str(item[0])}"] = date(2018, 1, 1)
            elif milestones_regex5.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2019, 1, 1)
                ws[f"C{str(item[0])}"] = date(2019, 1, 1)
                ws[f"D{str(item[0])}"] = date(2019, 1, 1)
            elif milestones_regex6.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2012, 1, 1)
                ws[f"C{str(item[0])}"] = date(2012, 1, 1)
                ws[f"D{str(item[0])}"] = date(2012, 1, 1)

            else:
                ws[f"B{str(item[0])}"] = " ".join([ix.upper(), "1"])
                ws[f"C{str(item[0])}"] = " ".join([ix.upper(), "2"])
                ws[f"D{str(item[0])}"] = " ".join([ix.upper(), "3"])
    wb.save(output_file)
    return output_file


@pytest.fixture(scope='session')
def previous_quarter_master():
    """
    This is a replica of the master() fixture above, but we're changing some
    values to simulate an earlier master than needs to be compared against.

    The values we're amending are the three values for "Working Contact Name",
    which appear in cells B11, C11, D11.
    :return: output_file
    """
    # regexes
    r'(Assurance MM1 .+$|Approval MM1 .+$)'
    milestones_regex1 = re.compile(r'(Assurance MM1 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM1 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex2 = re.compile(r'(Assurance MM2 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM2 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex3 = re.compile(r'(Assurance MM\d+ (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM\d+ (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex4 = re.compile(r'(Assurance MM3 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM3 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex5 = re.compile(r'(Assurance MM4 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM4 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex6 = re.compile(r'(Assurance MM5 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM5 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')

    wb = Workbook()
    output_file = "/".join([OUTPUT_DIR, 'previous_quarter_master.xlsx'])
    ws = wb.active
    ws.title = "Previous Master for Testing"
    for item in enumerate(real_datamap_data.split('\n')):
        if not item[0] == 0 and not item[1] == "":
            g = split_datamap_line(item)
            next(g)
            ix = next(g).split(',')[0]
            ws[f"A{str(item[0])}"] = ix
            if item[1].startswith('Date'):  # testing for date objects
                ws[f"B{str(item[0])}"] = date(2012, 6, 20)
                ws[f"C{str(item[0])}"] = date(2012, 6, 20)
                ws[f"D{str(item[0])}"] = date(2012, 6, 20)
            elif item[1].startswith('SRO Tenure'):  # testing for date objects
                ws[f"B{str(item[0])}"] = date(2017, 8, 10)
                ws[f"C{str(item[0])}"] = date(2017, 8, 10)
                ws[f"D{str(item[0])}"] = date(2017, 8, 10)
            elif item[1].startswith('Total Forecast'):
                ws[f"B{str(item[0])}"] = 32.3333
                ws[f"C{str(item[0])}"] = 35.2322
                ws[f"D{str(item[0])}"] = 23.2
            elif item[1].startswith('BICC approval point'):
                ws[f"B{str(item[0])}"] = "Strategic Outline Case"
                ws[f"C{str(item[0])}"] = "Outline Business Case"
                ws[f"D{str(item[0])}"] = "Full Business Case"
            elif item[1].startswith('Project MM20 Forecast - Actual'):
                ws[f"B{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"C{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"D{str(item[0])}"] = datetime(2016, 1, 1)
            elif item[1].startswith('Departmental DCA'):
                ws[f"B{str(item[0])}"] = "Amber"
                ws[f"C{str(item[0])}"] = "Red"
                ws[f"D{str(item[0])}"] = "Green"
            elif item[1].startswith('SRO Finance confidence'):
                ws[f"B{str(item[0])}"] = "Amber/Red"
                ws[f"C{str(item[0])}"] = "Amber/Red"
                ws[f"D{str(item[0])}"] = "Amber"
            elif item[1].startswith('SRO Benefits RAG'):
                ws[f"B{str(item[0])}"] = "Red"
                ws[f"C{str(item[0])}"] = "Green"
                ws[f"D{str(item[0])}"] = "Amber/Green"


            # Here we are starting a block of dates. We need these to be able to test
            # the default swimlane_milstones analyser
            # we're giving these ones some variety so they can be tested as the default
            # swimlane_milestones analyser
            elif milestones_regex1.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2015, 1, 1)
                ws[f"C{str(item[0])}"] = date(2015, 1, 1)
                ws[f"D{str(item[0])}"] = date(2015, 1, 1)
            elif milestones_regex2.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2019, 1, 1)
                ws[f"C{str(item[0])}"] = date(2019, 1, 1)
                ws[f"D{str(item[0])}"] = date(2019, 1, 1)
            elif milestones_regex3.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2020, 1, 1)
                ws[f"C{str(item[0])}"] = date(2020, 1, 1)
                ws[f"D{str(item[0])}"] = date(2020, 1, 1)
            elif milestones_regex4.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2018, 1, 1)
                ws[f"C{str(item[0])}"] = date(2018, 1, 1)
                ws[f"D{str(item[0])}"] = date(2018, 1, 1)
            elif milestones_regex5.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2019, 1, 1)
                ws[f"C{str(item[0])}"] = date(2019, 1, 1)
                ws[f"D{str(item[0])}"] = date(2019, 1, 1)
            elif milestones_regex6.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2012, 1, 1)
                ws[f"C{str(item[0])}"] = date(2012, 1, 1)
                ws[f"D{str(item[0])}"] = date(2012, 1, 1)

            else:
                ws[f"B{str(item[0])}"] = " ".join([ix.upper(), "1"])
                ws[f"C{str(item[0])}"] = " ".join([ix.upper(), "2"])
                ws[f"D{str(item[0])}"] = " ".join([ix.upper(), "3"])

    # FOR COMPARISON TESTS

    # here we amend the three string cells...
    ws['B11'].value = ' '.join([ws['B11'].value, 'AMENDED'])
    ws['C11'].value = ' '.join([ws['B11'].value, 'AMENDED'])
    ws['D11'].value = ' '.join([ws['B11'].value, 'AMENDED'])

    # here we amend a single date cells...
    # this is for "SRO Tenure Start Date"
    ws['B13'].value = date(2017, 3, 1)

    # now setting an later date for "SRO Tenure End Date"
    # also now for PROJECT/PROGRAMME NAME 2
    ws['C14'].value = date(2019, 6, 6)


    wb.save(output_file)
    return output_file


@pytest.fixture(scope='session')
def master_one_extra_proj():
    """
    This is master file created for the purpose of using a base for bcompiler -a, which
    populates all the returns. It simply takes the field name from the datamap and
    puts it in upper case and appends a digit (1, 2 or 3 because we're only simulating
    a master with 4 projects here.
    :return: output_file
    """
    # regexes
    r'(Assurance MM1 .+$|Approval MM1 .+$)'
    milestones_regex1 = re.compile(r'(Assurance MM1 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM1 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex2 = re.compile(r'(Assurance MM2 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM2 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex4 = re.compile(r'(Assurance MM3 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM3 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex5 = re.compile(r'(Assurance MM4 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM4 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex6 = re.compile(r'(Assurance MM5 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM5 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex3 = re.compile(r'(Assurance MM\d+ (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM\d+ (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')

    wb = Workbook()
    output_file = "/".join([OUTPUT_DIR, 'master_extra_project.xlsx'])
    ws = wb.active
    ws.title = "Master for Testing"
    for item in enumerate(real_datamap_data.split('\n')):
        if not item[0] == 0 and not item[1] == "":
            g = split_datamap_line(item)
            next(g)
            ix = next(g).split(',')[0]
            ws[f"A{str(item[0])}"] = ix
            if item[1].startswith('Date'):  # testing for date objects
                ws[f"B{str(item[0])}"] = date(2017, 6, 20)
                ws[f"C{str(item[0])}"] = date(2017, 6, 20)
                ws[f"D{str(item[0])}"] = date(2017, 6, 20)
                ws[f"E{str(item[0])}"] = date(2017, 6, 20)
            elif item[1].startswith('SRO Tenure'):  # testing for date objects
                ws[f"B{str(item[0])}"] = date(2017, 8, 10)
                ws[f"C{str(item[0])}"] = date(2017, 8, 10)
                ws[f"D{str(item[0])}"] = date(2017, 8, 10)
                ws[f"E{str(item[0])}"] = date(2017, 8, 10)
            elif item[1].startswith('Total Forecast'):
                ws[f"B{str(item[0])}"] = 32.3333
                ws[f"C{str(item[0])}"] = 35.2322
                ws[f"D{str(item[0])}"] = 23.2
                ws[f"E{str(item[0])}"] = 23.2
            elif item[1].startswith('BICC approval point'):
                ws[f"B{str(item[0])}"] = "Strategic Outline Case"
                ws[f"C{str(item[0])}"] = "Outline Business Case"
                ws[f"D{str(item[0])}"] = "Full Business Case"
                ws[f"E{str(item[0])}"] = "Full Business Case"
            elif item[1].startswith('Project MM20 Forecast - Actual'):
                ws[f"B{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"C{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"D{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"E{str(item[0])}"] = datetime(2016, 1, 1)
            elif item[1].startswith('Departmental DCA'):
                ws[f"B{str(item[0])}"] = "Amber/Red"
                ws[f"C{str(item[0])}"] = "Red"
                ws[f"D{str(item[0])}"] = "Green"
                ws[f"E{str(item[0])}"] = "Green"
            elif item[1].startswith('SRO Finance confidence'):
                ws[f"B{str(item[0])}"] = "Amber/Red"
                ws[f"C{str(item[0])}"] = "Amber/Red"
                ws[f"D{str(item[0])}"] = "Amber"
                ws[f"E{str(item[0])}"] = "Amber"
            elif item[1].startswith('SRO Benefits RAG'):
                ws[f"B{str(item[0])}"] = "Red"
                ws[f"C{str(item[0])}"] = "Green"
                ws[f"D{str(item[0])}"] = "Amber/Green"
                ws[f"E{str(item[0])}"] = "Amber/Green"
            elif item[1].startswith('GMPP - IPA DCA'):
                ws[f"B{str(item[0])}"] = "Amber"
                ws[f"C{str(item[0])}"] = "Amber"
                ws[f"D{str(item[0])}"] = "Amber/Green"
                ws[f"E{str(item[0])}"] = "Amber/Green"


            # Here we are starting a block of dates. We need these to be able to test
            # the default swimlane_milstones analyser
            # we're giving these ones some variety so they can be tested as the default
            # swimlane_milestones analyser
            elif milestones_regex1.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2015, 1, 1)
                ws[f"C{str(item[0])}"] = date(2015, 1, 1)
                ws[f"D{str(item[0])}"] = date(2015, 1, 1)
                ws[f"E{str(item[0])}"] = date(2015, 1, 1)
            elif milestones_regex2.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2019, 1, 1)
                ws[f"C{str(item[0])}"] = date(2019, 1, 1)
                ws[f"D{str(item[0])}"] = date(2019, 1, 1)
                ws[f"E{str(item[0])}"] = date(2019, 1, 1)
            elif milestones_regex3.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2020, 1, 1)
                ws[f"C{str(item[0])}"] = date(2020, 1, 1)
                ws[f"D{str(item[0])}"] = date(2020, 1, 1)
                ws[f"E{str(item[0])}"] = date(2020, 1, 1)
            elif milestones_regex4.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2018, 1, 1)
                ws[f"C{str(item[0])}"] = date(2018, 1, 1)
                ws[f"D{str(item[0])}"] = date(2018, 1, 1)
                ws[f"E{str(item[0])}"] = date(2018, 1, 1)
            elif milestones_regex5.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2019, 1, 1)
                ws[f"C{str(item[0])}"] = date(2019, 1, 1)
                ws[f"D{str(item[0])}"] = date(2019, 1, 1)
                ws[f"E{str(item[0])}"] = date(2019, 1, 1)
            elif milestones_regex6.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2012, 1, 1)
                ws[f"C{str(item[0])}"] = date(2012, 1, 1)
                ws[f"D{str(item[0])}"] = date(2012, 1, 1)
                ws[f"E{str(item[0])}"] = date(2012, 1, 1)

            else:
                ws[f"B{str(item[0])}"] = " ".join([ix.upper(), "1"])
                ws[f"C{str(item[0])}"] = " ".join([ix.upper(), "2"])
                ws[f"D{str(item[0])}"] = " ".join([ix.upper(), "3"])
                ws[f"E{str(item[0])}"] = " ".join([ix.upper(), "4"])
    wb.save(output_file)
    return output_file


@pytest.fixture(scope='session')
def master_with_quarter_year_in_filename():
    """
    This is a replica of the master() fixture above, but we're chanching the
    name of the output file to be of the form master_1_2017.xlsx, as this
    is what the financial analyser is going to need.

    :return: output_file
    """
    # regexes
    r'(Assurance MM1 .+$|Approval MM1 .+$)'
    milestones_regex1 = re.compile(r'(Assurance MM1 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM1 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex2 = re.compile(r'(Assurance MM2 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM2 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex3 = re.compile(r'(Assurance MM\d+ (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM\d+ (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$)')
    milestones_regex4 = re.compile(r'(Assurance MM3 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM3 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex5 = re.compile(r'(Assurance MM4 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM4 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')
    milestones_regex6 = re.compile(r'(Assurance MM5 (?!Notes)(?!Milestone Type)(?!Type)(?!DCA).+$|Approval MM5 (?!Notes).+$)(?!Milestone Type)(?!Type)(?!DCA)')

    milestones_regex7 = re.compile(r'Project MM\d+ Forecast - Actual')

    wb = Workbook()
    output_file = "/".join([OUTPUT_DIR, 'previous_quarter_master_1_2017.xlsx'])
    ws = wb.active
    ws.title = "Previous Master for Testing"
    for item in enumerate(real_datamap_data.split('\n')):
        if not item[0] == 0 and not item[1] == "":
            g = split_datamap_line(item)
            next(g)
            ix = next(g).split(',')[0]
            ws[f"A{str(item[0])}"] = ix
            if item[1].startswith('Date'):  # testing for date objects
                ws[f"B{str(item[0])}"] = date(2017, 6, 20)
                ws[f"C{str(item[0])}"] = date(2017, 6, 20)
                ws[f"D{str(item[0])}"] = date(2017, 6, 20)
            elif item[1].startswith('SRO Tenure'):  # testing for date objects
                ws[f"B{str(item[0])}"] = date(2017, 8, 10)
                ws[f"C{str(item[0])}"] = date(2017, 8, 10)
                ws[f"D{str(item[0])}"] = date(2017, 8, 10)
            elif item[1].startswith('Total Forecast'):
                ws[f"B{str(item[0])}"] = 32.3333
                ws[f"C{str(item[0])}"] = 35.2322
                ws[f"D{str(item[0])}"] = 23.2
            elif item[1].startswith('BICC approval point'):
                ws[f"B{str(item[0])}"] = "Strategic Outline Case"
                ws[f"C{str(item[0])}"] = "Outline Business Case"
                ws[f"D{str(item[0])}"] = "Full Business Case"
            elif item[1].startswith('Project MM20 Forecast - Actual'):
                ws[f"B{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"C{str(item[0])}"] = datetime(2016, 1, 1)
                ws[f"D{str(item[0])}"] = datetime(2016, 1, 1)
            elif item[1].startswith('Departmental DCA'):
                ws[f"B{str(item[0])}"] = "Amber"
                ws[f"C{str(item[0])}"] = "Red"
                ws[f"D{str(item[0])}"] = "Green"
            elif item[1].startswith('SRO Finance confidence'):
                ws[f"B{str(item[0])}"] = "Amber/Red"
                ws[f"C{str(item[0])}"] = "Amber/Red"
                ws[f"D{str(item[0])}"] = "Amber"
            elif item[1].startswith('SRO Benefits RAG'):
                ws[f"B{str(item[0])}"] = "Red"
                ws[f"C{str(item[0])}"] = "Green"
                ws[f"D{str(item[0])}"] = "Amber/Green"


            # Here we are starting a block of dates. We need these to be able to test
            # the default swimlane_milstones analyser
            # we're giving these ones some variety so they can be tested as the default
            # swimlane_milestones analyser
            elif milestones_regex1.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2015, 1, 1)
                ws[f"C{str(item[0])}"] = date(2015, 1, 1)
                ws[f"D{str(item[0])}"] = date(2015, 1, 1)
            elif milestones_regex2.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2019, 1, 1)
                ws[f"C{str(item[0])}"] = date(2019, 1, 1)
                ws[f"D{str(item[0])}"] = date(2019, 1, 1)
            elif milestones_regex3.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2020, 1, 1)
                ws[f"C{str(item[0])}"] = date(2020, 1, 1)
                ws[f"D{str(item[0])}"] = date(2020, 1, 1)
            elif milestones_regex4.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2018, 1, 1)
                ws[f"C{str(item[0])}"] = date(2018, 1, 1)
                ws[f"D{str(item[0])}"] = date(2018, 1, 1)
            elif milestones_regex5.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2019, 1, 1)
                ws[f"C{str(item[0])}"] = date(2019, 1, 1)
                ws[f"D{str(item[0])}"] = date(2019, 1, 1)
            elif milestones_regex6.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2012, 1, 1)
                ws[f"C{str(item[0])}"] = date(2012, 1, 1)
                ws[f"D{str(item[0])}"] = date(2012, 1, 1)
            elif milestones_regex7.match(item[1]):
                ws[f"B{str(item[0])}"] = date(2013, 1, 1)
                ws[f"C{str(item[0])}"] = date(2013, 1, 1)
                ws[f"D{str(item[0])}"] = date(2013, 1, 1)


            else:
                ws[f"B{str(item[0])}"] = " ".join([ix.upper(), "1"])
                ws[f"C{str(item[0])}"] = " ".join([ix.upper(), "2"])
                ws[f"D{str(item[0])}"] = " ".join([ix.upper(), "3"])

    # FOR COMPARISON TESTS

    # here we amend the three string cells...
    ws['B11'].value = ' '.join([ws['B11'].value, 'AMENDED'])
    ws['C11'].value = ' '.join([ws['B11'].value, 'AMENDED'])
    ws['D11'].value = ' '.join([ws['B11'].value, 'AMENDED'])

    # here we amend a single date cells...
    # this is for "SRO Tenure Start Date"
    ws['B13'].value = date(2017, 3, 1)

    # now setting an later date for "SRO Tenure End Date"
    # also now for PROJECT/PROGRAMME NAME 2
    ws['C14'].value = date(2019, 6, 6)


    wb.save(output_file)
    return output_file
