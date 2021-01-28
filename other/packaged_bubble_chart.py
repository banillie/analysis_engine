"""From search matplotlib package bubble chart search"""

import numpy as np
import matplotlib.pyplot as plt
from analysis_engine.data import Master, open_pickle_file, root_path, convert_rag_text

colour_dict = {
    "A": "00fce553",
    "A/G": "00a5b700",
    "A/R": "00f97b31",
    "R": "00fc2525",
    "G": "0017960c",
    "": "#808080"  # Gray if missing
}


def get_cost(master: Master, project_names: list, output_list: list):
    for x in range(len(project_names)):
        t = master.master_data[0].data[project_names[x]]["Total Forecast"]
        output_list.append(t)

    return output_list


def get_colours(master: Master, project_names: list, output_list: list):
    for x in range(len(project_names)):
        c = convert_rag_text(master.master_data[0].data[project_names[x]]["Departmental DCA"])
        output_list.append(colour_dict[c])

    return output_list


def browser_marker_share(master: Master):
    return {
        "browsers": master.current_projects,
        "market_share": get_cost(master, master.current_projects, []),
        "colour": get_colours(master, master.current_projects, [])
    }


class BubbleChart:
    def __init__(self, area, bubble_spacing=0):
        """
        Setup for bubble collapse.

        @param area: array-like. Area of the bubbles.
        @param bubble_spacing: float, default:0. Minimal spacing between bubbles after collapsing.

        @note
        If "area" is sorted, the results might look weird.
        """
        area = np.asarray(area)
        r = np.sqrt(area / np.pi)

        self.bubble_spacing = bubble_spacing
        self.bubbles = np.ones((len(area), 4))
        self.bubbles[:, 2] = r
        self.bubbles[:, 3] = area
        self.maxstep = 2 * self.bubbles[:, 2].max() + self.bubble_spacing
        self.step_dist = self.maxstep / 2

        # calculate initial grid layout for bubbles
        length = np.ceil(np.sqrt(len(self.bubbles)))
        grid = np.arrange(length) * self.maxstep  # arrange might cause trouble
        gx, gy = np.meshgrid(grid, grid)
        self.bubbles[:, 0] = gx.flatten()[:len(self.bubbles)]

m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
test = browser_marker_share(m)
