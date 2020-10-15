"""
New code for compiling individual project reports.
"""
import os

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from data_mgmt.oldegg_functions import convert_rag_text, cell_colouring, make_rows_bold, set_col_widths, \
    compare_text_newandold

from data_mgmt.data import root_path, SRO_conf_key_list, project_cost_profile_graph


# project_info = project_data_from_master(root_path / "core_data/project_info.xlsx", 1, 2099)

def open_word_doc(wd_path: str):
    """Function stores an empty word doc as a variable"""
    return Document(wd_path)


def wd_heading(doc, project_info, project):
    """Function adds header to word doc"""
    font = doc.styles['Normal'].font
    font.name = 'Arial'
    font.size = Pt(12)

    heading = str(project_info.data[project]['Abbreviations'])
    intro = doc.add_heading(str(heading), 0)
    intro.alignment = 1
    intro.bold = True


def key_contacts(doc, master, project):
    """Function adds key contact details"""
    sro_name = master.master_data[0].data[project]['Senior Responsible Owner (SRO)']
    if sro_name is None:
        sro_name = 'tbc'

    sro_email = master.master_data[0].data[project]['Senior Responsible Owner (SRO) - Email']
    if sro_email is None:
        sro_email = 'email: tbc'

    sro_phone = master.master_data[0].data[project]['SRO Phone No.']
    if sro_phone == None:
        sro_phone = 'phone number: tbc'

    doc.add_paragraph('SRO: ' + str(sro_name) + ', ' + str(sro_email) + ', ' + str(sro_phone))

    pd_name = master.master_data[0].data[project]['Project Director (PD)']
    if pd_name is None:
        pd_name = 'TBC'

    pd_email = master.master_data[0].data[project]['Project Director (PD) - Email']
    if pd_email is None:
        pd_email = 'email: tbc'

    pd_phone = master.master_data[0].data[project]['PD Phone No.']
    if pd_phone is None:
        pd_phone = 'TBC'

    doc.add_paragraph('PD: ' + str(pd_name) + ', ' + str(pd_email) + ', ' + str(pd_phone))

    contact_name = master.master_data[0].data[project]['Working Contact Name']
    if contact_name is None:
        contact_name = 'TBC'

    contact_email = master.master_data[0].data[project]['Working Contact Email']
    if contact_email is None:
        contact_email = 'email: tbc'

    contact_phone = master.master_data[0].data[project]['Working Contact Telephone']
    if contact_phone is None:
        contact_phone = 'TBC'

    doc.add_paragraph('PfM reporting lead: ' + str(contact_name) + ', ' + str(contact_email)
                      + ', ' + str(contact_phone))


def dca_table(doc, master, project):
    """Creates SRO confidence table"""
    table = doc.add_table(rows=1, cols=5)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Delivery confidence'
    hdr_cells[1].text = 'This quarter'
    hdr_cells[2].text = str(master.master_data[1].quarter)
    hdr_cells[3].text = str(master.master_data[2].quarter)
    hdr_cells[4].text = str(master.master_data[3].quarter)

    for x, dca_key in enumerate(SRO_conf_key_list):
        row_cells = table.add_row().cells
        row_cells[0].text = dca_key
        for i, m in enumerate(master.master_data):
            try:
                rating = convert_rag_text(m.data[project][dca_key])
                row_cells[i + 1].text = rating
                cell_colouring(row_cells[i + 1], rating)
            except (KeyError, TypeError):
                row_cells[i + 1].text = "N/A"

    table.style = 'Table Grid'
    make_rows_bold([table.rows[0]])  # makes top of table bold.
    # make_columns_bold([table.columns[0]]) #right cells in table bold
    column_widths = (Cm(3.9), Cm(2.9), Cm(2.9), Cm(2.9), Cm(2.9))
    set_col_widths(table, column_widths)


def dca_narratives(doc, master, project):
    """Places all narratives into document and checks for differences between
    current and last quarter"""

    doc.add_paragraph()
    p = doc.add_paragraph()
    text = '*Red text highlights changes in narratives from last quarter'
    p.add_run(text).font.color.rgb = RGBColor(255, 0, 0)

    headings_list = ['SRO delivery confidence narrative',
                     'Financial cost narrative',
                     'Financial comparison with last quarter',
                     'Financial comparison with baseline',
                     'Benefits Narrative',
                     'Benefits comparison with last quarter',
                     'Benefits comparison with baseline',
                     'Milestone narrative']

    narrative_keys_list = ['Departmental DCA Narrative',
                           'Project Costs Narrative',
                           'Cost comparison with last quarters cost narrative',
                           'Cost comparison within this quarters cost narrative',
                           'Benefits Narrative',
                           'Ben comparison with last quarters cost - narrative',
                           'Ben comparison within this quarters cost - narrative',
                           'Milestone Commentary']

    for x in range(len(headings_list)):
        doc.add_paragraph().add_run(str(headings_list[x])).bold = True
        text_one = str(master.master_data[0].data[project][narrative_keys_list[x]])
        try:
            text_two = str(master.master_data[1].data[project][narrative_keys_list[x]])
        except KeyError:
            text_two = text_one

        # There are two options here for comparing text. Have left this for now.
        # compare_text_showall(dca_a, dca_b, doc)
        compare_text_newandold(text_one, text_two, doc)


def year_cost_profile_chart(doc, master):
    """Places line graph cost profile into word document"""

    new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)  # new page
    # change to landscape
    new_width, new_height = new_section.page_height, new_section.page_width
    new_section.orientation = WD_ORIENT.LANDSCAPE
    new_section.page_width = new_width
    new_section.page_height = new_height

    fig = project_cost_profile_graph(master)

    # Size and shape of figure.
    fig.canvas.draw()
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # for title

    # Place fig in word doc.
    fig.savefig('cost_profile.png')
    doc.add_picture('cost_profile.png', width=Inches(8))  # to place nicely in doc
    os.remove('cost_profile.png')
