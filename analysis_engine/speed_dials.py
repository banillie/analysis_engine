import numpy as np
from docx.shared import Inches
from matplotlib import pyplot as plt
from matplotlib.patches import Wedge, Circle
from typing import List

from analysis_engine.dictionaries import DCA_RATING_SCORES, FONT_TYPE
from analysis_engine.colouring import FACE_COLOUR, COLOUR_DICT
from analysis_engine.render_utils import put_matplotlib_fig_into_word


# for speed_dial analysis_engine code taken from
# http://nicolasfauchereau.github.io/climatecode/posts/drawing-a-gauge-with-matplotlib/
def degree_range(n):
    start = np.linspace(-30, 210, n + 1, endpoint=True)[0:-1]
    end = np.linspace(-30, 210, n + 1, endpoint=True)[1::]
    mid_points = start + ((end - start) / 2.0)
    return np.c_[start, end], mid_points


def rot_text(ang):
    rotation = np.degrees(np.radians(ang) * np.pi / np.pi - np.radians(90))
    return rotation


def gauge(
    labels: List[str],
    total: str,
    arrow_one: float,
    arrow_two: float,
    up: str,
    down: str,
    title: str,
):
    no = len(labels)
    fig, ax = plt.subplots(facecolor=FACE_COLOUR)
    fig.set_size_inches(18.5, 10.5)
    ax.set_facecolor(FACE_COLOUR)  # TBC if face colour is required
    ang_range, mid_points = degree_range(no)
    if no == 3:
        colours = [
            COLOUR_DICT["R"],
            COLOUR_DICT["A"],
            COLOUR_DICT["G"],
        ]
    else:
        colours = [
            COLOUR_DICT["R"],
            COLOUR_DICT["A/R"],
            COLOUR_DICT["A"],
            COLOUR_DICT["A/G"],
            COLOUR_DICT["G"],
        ]

    patches = []
    for ang, c in zip(ang_range, colours):
        patches.append(
            Wedge(
                (0.0, 0.0),
                0.4,
                *ang,
                width=0.15,
                facecolor=c,
                lw=2,
                ec="black",
                zorder=2,
            )
        )

    [ax.add_patch(p) for p in patches]

    for mid, lab in zip(mid_points, reversed(labels)):
        ax.text(
            0.325 * np.cos(np.radians(mid)),
            0.325 * np.sin(np.radians(mid)),
            lab,
            horizontalalignment="center",
            verticalalignment="center",
            fontsize=35,
            fontweight="bold",
            fontname=FONT_TYPE,
            rotation=rot_text(mid),
            zorder=4,
        )

    ## title
    plt.suptitle(
        title.upper() + " Confidence",
        fontweight="bold",
        fontsize=50,
        fontname=FONT_TYPE,
    )

    ax.text(
        0,
        -0.1,
        total,
        horizontalalignment="center",
        verticalalignment="center",
        fontsize=60,
        fontweight="bold",
        fontname=FONT_TYPE,
        zorder=1,
    )

    ax.text(
        0,
        -0.17,
        "SRO ratings",
        horizontalalignment="center",
        verticalalignment="center",
        fontsize=45,
        fontweight="bold",
        fontname=FONT_TYPE,
        zorder=1,
    )

    def get_arrow_point(score: float):
        return (240 * score) - 120

    ax.annotate(
        "",
        xy=(
            (0.275 * np.sin(np.radians(get_arrow_point(arrow_two)))),
            (0.275 * np.cos(np.radians(get_arrow_point(arrow_two)))),
        ),
        xytext=(0, 0),
        arrowprops=dict(
            arrowstyle="wedge", color="grey", fill=True, alpha=0.5, linewidth=20  #
        ),
    )

    ax.annotate(
        "",
        xy=(
            (0.275 * np.sin(np.radians(get_arrow_point(arrow_one)))),
            (0.275 * np.cos(np.radians(get_arrow_point(arrow_one)))),
        ),
        xytext=(0, 0),
        arrowprops=dict(arrowstyle="wedge", linewidth=20),
    )

    ax.add_patch(Circle((0, 0), radius=0.02, facecolor="k"))
    ax.add_patch(Circle((0, 0), radius=0.01, facecolor="w", zorder=11))

    ## arrows around the dial
    if arrow_one != arrow_two:  # only done if two quarters data available.
        plt.annotate(
            "",
            xy=(-0.2, 0.4),
            xycoords="data",
            xytext=(-0.4, 0.2),
            textcoords="data",
            arrowprops=dict(
                arrowstyle="->", connectionstyle="arc3, rad=-0.3", linewidth=4
            ),
        )
        ax.text(-0.35, 0.35, down, fontsize=30, fontname=FONT_TYPE)

        ax.text(0.35, 0.35, up, fontsize=30, fontname=FONT_TYPE)

        plt.annotate(
            "",
            xy=(0.4, 0.2),
            xycoords="data",
            xytext=(0.2, 0.4),
            textcoords="data",
            arrowprops=dict(
                arrowstyle="<-", connectionstyle="arc3, rad=-0.3", linewidth=4
            ),
        )

    plt.axis("scaled")
    plt.axis("off")

    return fig


def delete_from_c_count(count_list: list, to_remove: list):
    """
    As dca rag ratings have been standardised via DCA_RATINGS_SCORES a small helper
    function is require to get the c_count list and rate list used in function
    build_speed_dials to match. This is necessary as cdg report uses five rag ratings
    while ipdc uses three.
    """
    for x in to_remove:
        del count_list[x]
    return count_list


def build_speed_dials(dca_data, doc):
    for conf_type in dca_data.dca_count[dca_data.quarters[0]]:
        c_count = []
        l_count = []
        for colour in DCA_RATING_SCORES.keys():
            c_no = dca_data.dca_count[dca_data.quarters[0]][conf_type][colour]["count"]
            c_count.append(c_no)
            try:  # to capture reporting process with only one quarters data
                l_no = dca_data.dca_count[dca_data.quarters[1]][conf_type][colour][
                    "count"
                ]
                l_count.append(l_no)
            except KeyError:
                l_count.append(c_no)
            except IndexError:
                pass
        c_total = dca_data.dca_count[dca_data.quarters[0]][conf_type]["Total"]["count"]

        up = 0
        down = 0

        if bool(dca_data.dca_changes):  # empty dictionaries evaluation to false
            for p in dca_data.dca_changes[conf_type]:
                change = dca_data.dca_changes[conf_type][p]["Change"]
                if change == "Up":
                    up += 1
                if change == "Down":
                    down += 1

        if dca_data.kwargs["rag_number"] == "3":
            rate = [0, 0.5, 1]
            c_count = delete_from_c_count(c_count, [1, 2, 3])
            l_count = delete_from_c_count(l_count, [1, 2, 3])
        if dca_data.kwargs["rag_number"] == "5":
            rate = [0, 0.25, 0.5, 0.75, 1]
            c_count = delete_from_c_count(c_count, [-1])
            l_count = delete_from_c_count(l_count, [-1])
        dial_one = np.average(rate, weights=c_count)
        try:
            dial_two = np.average(rate, weights=l_count)
        except TypeError:  # no previous quarter data
            dial_two = dial_one

        graph = gauge(
            c_count,
            str(c_total),
            dial_one,
            dial_two,
            str(up),
            str(down),
            title=conf_type,
        )

        put_matplotlib_fig_into_word(doc, graph, width=Inches(8))
