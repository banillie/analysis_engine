import os

from matplotlib import pyplot as plt
from docx import Document
from docx.shared import Inches
from pdf2image import convert_from_path



def put_matplotlib_fig_into_word(
        doc: Document, fig: plt.figure or plt, **kwargs
) -> None:
    """Does rendering of matplotlib graph into word. Best method I could find for
    maintain high quality render output it to firstly save as pdf and then convert
    to jpeg!"""
    # Place fig in word doc.
    fig.savefig("fig.pdf")
    # fig.savefig("cost_profile.png", dpi=300)
    # fig.savefig("cost_profile.png", bbox_inches="tight")
    page = convert_from_path("fig.pdf", 500)
    page[0].save("fig.jpeg", "JPEG")
    if "size" in kwargs:
        s = kwargs["size"]
        doc.add_picture("fig.jpeg", width=Inches(s))
    else:
        doc.add_picture("fig.jpeg", width=Inches(8))  # to place nicely in doc
    os.remove("fig.jpeg")
    os.remove("fig.pdf")
    plt.close()  # automatically closes figure so don't need to do manually.

