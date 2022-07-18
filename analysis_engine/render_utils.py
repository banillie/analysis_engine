import os
import re

from typing import TextIO, Union, Tuple, List
from matplotlib import pyplot as plt
from docx import Document
from docx.shared import Inches
from pdf2image import convert_from_path
from openpyxl import Workbook, load_workbook
from textwrap import wrap

from analysis_engine.error_msgs import logger


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


def put_matplotlib_fig_into_word(doc, fig: plt.figure or plt, **kwargs) -> None:
    """Does rendering of matplotlib graph into word. Best method I could find for
    maintain high quality render output it to firstly save as pdf and then convert
    to jpeg!
    kwargs can be width=Inches(int) or transparent=False
    """
    fig.savefig("fig.pdf")
    page = convert_from_path("fig.pdf", 500)
    page[0].save("fig.jpeg", "JPEG")
    doc.add_picture("fig.jpeg", **kwargs)  # to place nicely in doc
    # doc.add_picture("fig.jpeg")
    os.remove("fig.jpeg")
    os.remove("fig.pdf")
    plt.close()  # automatically closes figure so don't need to do manually.


def make_file_friendly(quarter_str: str) -> str:
    """Converts datamaps.api project_data_from_master quarter data into a string to use when
    saving output files. Courtesy of M Lemon."""
    regex = r"Q(\d) (\d+)\/(\d+)"
    return re.sub(regex, r"Q\1_\2_\3", quarter_str)


FIGURE_STYLE = {1: "half_horizontal", 2: "full_horizontal"}


def set_figure_size(graph_type: str) -> Tuple[int, int]:
    if graph_type == "half_horizontal":
        return 11.69, 5.10
    if graph_type == "full_horizontal":
        return 11.69, 8.20


def set_fig_size(kwargs, fig: plt.figure) -> plt.figure:
    if "fig_size" in kwargs:
        fig.set_size_inches(set_figure_size(kwargs["fig_size"]))
    else:
        fig.set_size_inches(set_figure_size(FIGURE_STYLE[2]))

    return fig


def get_chart_title(
    **c_kwargs,  # chart kwargs
) -> str:
    if "title" in c_kwargs:
        return c_kwargs["title"]
    else:
        logger.info("Please provide a title for this chart using --title.")
        return None


# helper function for milestone chart
def handle_long_keys(project_name: str) -> str:
    if len(project_name) >= 25:
        l = project_name.split()
        word_count = len(l)
        if word_count == 1:
            return project_name
        if word_count == 2:
            l.insert(1, '\n')
        if word_count >= 4:
            # if len(" ".join(l[3])) <= 10:
            #     l.insert(3, '\n')
            #     l.insert(6, '\n')
            #     l.insert(9, '\n')
            # else:
            # l.insert(2, '\n')
        # if word_count >= 4:   # to improve.
                l.insert(3, '\n')
                # l.insert(5, '\n')
                # l.insert(8, '\n')
        new_str = " ".join(l)
        return re.sub('\s\\n\s', '\n', new_str)
    else:
        return project_name

