"""
Outputs analysis for speed_dials. Outputs are place in analysis_engine/output. They are:
- A word document titled dca_changes which specifies which project dca ratings have changed
- matplotlib speed dial?
"""

# from data_mgmt.data import (
#     Master,
#     root_path,
#     dca_changes_into_word,
#     get_project_information,
#     get_master_data,
#     get_word_doc,
#     DcaData,
# )
#
#
# def compile_speed_dials():
#     m = Master(get_master_data(), get_project_information())
#     dca = DcaData(m)
#     latest_quarter = str(m.master_data[0].quarter)
#     last_quarter = str(m.master_data[1].quarter)
#     dca.get_changes(latest_quarter, last_quarter)
#     word_doc = dca_changes_into_word(dca, get_word_doc())
#     word_doc.save(root_path / "output/dca_changes.docx")
#
#
# compile_speed_dials()


import os, sys
import matplotlib
from matplotlib import cm
from matplotlib import pyplot as plt
import numpy as np
from matplotlib.patches import Circle, Wedge, Rectangle
from docx.shared import Inches

from data_mgmt.data import open_word_doc, root_path


def degree_range(n):
    start = np.linspace(0,180,n+1, endpoint=True)[0:-1]
    end = np.linspace(0,180,n+1, endpoint=True)[1::]
    mid_points = start + ((end-start)/2.)
    return np.c_[start, end], mid_points


def rot_text(ang):
    rotation = np.degrees(np.radians(ang) * np.pi / np.pi - np.radians(90))
    return rotation


def gauge(labels=['LOW', 'MEDIUM', 'HIGH', 'VERY HIGH', 'EXTREME'], \
          colors='jet_r', arrow=1, arrow_two=2, title='', fname=False):
    """
    some sanity checks first

    """

    N = len(labels)

    if arrow > N:
        raise Exception("\n\nThe category ({}) is greated than \
        the length\nof the labels ({})".format(arrow, N))

    """
    if colors is a string, we assume it's a matplotlib colormap
    and we discretize in N discrete colors 
    """

    if isinstance(colors, str):
        cmap = cm.get_cmap(colors, N)
        cmap = cmap(np.arange(N))
        colors = cmap[::-1, :].tolist()
    if isinstance(colors, list):
        if len(colors) == N:
            colors = colors[::-1]
        else:
            raise Exception("\n\nnumber of colors {} not equal \
            to number of categories{}\n".format(len(colors), N))

    """
    begins the plotting
    """

    fig, ax = plt.subplots()

    ang_range, mid_points = degree_range(N)
    print(ang_range)
    print(mid_points)

    labels = labels[::-1]

    """
    plots the sectors and the arcs
    """
    patches = []
    for ang, c in zip(ang_range, colors):
        # sectors
        patches.append(Wedge((0., 0.), .4, *ang, facecolor='w', lw=2))
        # arcs
        patches.append(Wedge((0., 0.), .4, *ang, width=0.10, facecolor=c, lw=2, alpha=0.5))

    [ax.add_patch(p) for p in patches]

    """
    set the labels (e.g. 'LOW','MEDIUM',...)
    """

    for mid, lab in zip(mid_points, labels):
        ax.text(0.35 * np.cos(np.radians(mid)), 0.35 * np.sin(np.radians(mid)), lab, \
                horizontalalignment='center', verticalalignment='center', fontsize=14, \
                fontweight='bold', rotation=rot_text(mid))

    """
    set the bottom banner and the title
    """
    r = Rectangle((-0.4, -0.1), 0.8, 0.1, facecolor='w', lw=2)
    ax.add_patch(r)

    ax.text(0, -0.05, title, horizontalalignment='center', \
            verticalalignment='center', fontsize=22, fontweight='bold')

    """
    plots the arrow now
    """

    # pos = abs(arrow - N)
    pos = mid_points[abs(arrow - N)]
    print(pos)

    ax.arrow(0, 0, 0.225 * np.cos(np.radians(pos)), 0.225 * np.sin(np.radians(pos)), \
             width=0.04, head_width=0.09, head_length=0.05, fc='k', ec='k')

    pos_two = mid_points[abs(arrow_two - N)]

    ax.arrow(0, 0, 0.225 * np.cos(np.radians(pos_two)), 0.225 * np.sin(np.radians(pos_two)), \
             width=0.04, head_width=0.09, head_length=0.1, fc=None, ec='k')

    ax.add_patch(Circle((0, 0), radius=0.02, facecolor='k'))
    ax.add_patch(Circle((0, 0), radius=0.01, facecolor='w', zorder=11))

    """
    removes frame and ticks, and makes axis equal and tight
    """

    ax.set_frame_on(False)
    ax.axes.set_xticks([])
    ax.axes.set_yticks([])
    ax.axis('equal')
    plt.tight_layout()
    if fname:
        fig.savefig("temp_file.png", dpi=200)
        doc = open_word_doc(root_path / "input/summary_temp.docx")
        doc.add_picture("temp_file.png", width=Inches(8))
        doc.save(root_path / "output/speed_dial.docx")
        os.remove("temp_file.png")


gauge(labels=['R','A/R','A','A/G', 'G'], \
      colors=['#c00000','#e77200','#ffba00','#92a700', '#007d00'], arrow=3, arrow_two=2, title='DCA OVERALL', fname=True)

