import os

from typing import TextIO, Union

from matplotlib import pyplot as plt
from docx import Document
from docx.shared import Inches
from pdf2image import convert_from_path
from openpyxl import Workbook, load_workbook


def open_word_doc(wd_path: str) -> Document:
    """Function stores an empty word doc as a variable"""
    return Document(wd_path)


def get_input_doc(file_path: TextIO) -> Union[Workbook, Document, None]:
    """
    Returns blank documents in analysis_engine/input file used for saving outputs.
    Raises error and user message if files are not present
    """
    try:
        if str(file_path).endswith(".docx"):
            return open_word_doc(file_path)
        if str(file_path).endswith(".xlsx"):
            return load_workbook(file_path)
    except FileNotFoundError:
        base = os.path.basename(file_path)
        raise FileNotFoundError(
            str(base) + " document not present in input file. Stopping."
        )


def put_matplotlib_fig_into_word(doc, fig: plt.figure or plt) -> None:
    """Does rendering of matplotlib graph into word. Best method I could find for
    maintain high quality render output it to firstly save as pdf and then convert
    to jpeg!"""
    fig.savefig("fig.pdf")
    page = convert_from_path("fig.pdf", 500)
    page[0].save("fig.jpeg", "JPEG")
    # if "size" in kwargs:
    #     s = kwargs["size"]
    #     doc.add_picture("fig.jpeg", width=Inches(s))
    # else:
    doc.add_picture("fig.jpeg", width=Inches(8))  # to place nicely in doc
    os.remove("fig.jpeg")
    os.remove("fig.pdf")
    plt.close()  # automatically closes figure so don't need to do manually.


def make_file_friendly(quarter_str: str) -> str:
    """Converts datamaps.api project_data_from_master quarter data into a string to use when
    saving output files. Courtesy of M Lemon."""
    regex = r"Q(\d) (\d+)\/(\d+)"
    return re.sub(regex, r"Q\1_\2_\3", quarter_str)