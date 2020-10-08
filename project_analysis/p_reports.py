"""
New code for compiling individual project reports.
"""

from docx import Document

def open_word_doc(wd_path: str):
    """Function stores an empty word doc as a variable"""
    return Document(wd_path)


def wd_heading(doc):
    pass
