"""
New code for compiling individual project reports.
"""

from docx import Document
from datamaps.api import project_data_from_master
from docx.shared import Pt

from data_mgmt.data import root_path


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
    """creates SRO confidence table"""
