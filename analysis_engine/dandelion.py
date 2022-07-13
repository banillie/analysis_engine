import math
import datetime
from typing import List

from openpyxl import Workbook
from openpyxl.workbook import workbook

from matplotlib import pyplot as plt

from analysis_engine.costs import CostData
from analysis_engine.benefits import BenefitsData
from analysis_engine.resourcing import ResourceData
from analysis_engine.milestones import MilestoneData, get_milestone_date
from analysis_engine.segmentation import (
    get_iter_list,
    get_group,
    get_quarter_index,
    get_correct_p_data,
)
from analysis_engine.error_msgs import InputError, logger
from analysis_engine.colouring import COLOUR_DICT, FACE_COLOUR
from analysis_engine.render_utils import make_file_friendly

from analysis_engine.dictionaries import (
    RAG_RANKING_DICT_NUMBER,
    RAG_RANKING_DICT_COLOUR,
    BC_STAGE_DICT,
    DCA_KEYS,
    FONT_TYPE,
    STANDARDISE_DCA_KEYS,
    convert_rag_text,
    NEXT_STAGE_DICT,
)


def dandelion_project_text(number: int, project: str) -> str:
    total_len = len(str(int(number)))
    try:
        if total_len <= 3:
            round_total = int(round(number, -1))
            return "£" + str(round_total) + "m"
        if total_len == 4:
            round_total = int(round(number, -2))
            return "£" + str(round_total)[0] + "," + str(round_total)[1] + "bn"
        if total_len == 5:
            round_total = int(round(number, -2))
            return "£" + str(round_total)[:2] + "," + str(round_total)[2] + "bn"
        if total_len > 6:
            print(
                "Check total forecast and cost data reported by "
                + project
                + " total is £"
                + str(number)
                + "m"
            )
    except ValueError:
        print(
            "Check total forecast and cost data reported by "
            + project
            + " it is not reporting a number"
        )


def dandelion_number_text(number: int, **kwargs) -> str:
    total_len = len(str(int(number)))
    if "type" in kwargs:
        if kwargs["type"] in [
            "ps resource",
            "contract resource",
            "total resource",
            "funded resource",
        ]:
            return str(round(number, 1))
    try:
        if number == 0:
            if "none_handle" in kwargs:
                if kwargs["none_handle"] == "none":
                    return "£0m"
            else:
                return "£TBC"
        if total_len <= 2:
            # round_total = round(number, -1)
            return "£" + str(int(round(number))) + "m"
        if total_len == 3:
            round_total = round(number, -1)
            return "£" + str(int(round_total)) + "m"
        if total_len == 4:
            round_total = round(number, -2)
            if str(round_total)[1] != "0":
                return "£" + str(round_total)[0] + "," + str(round_total)[1] + "bn"
            else:
                return "£" + str(round_total)[0] + "bn"
        if total_len == 5:
            round_total = int(round(number, -2))
            if str(round_total)[2] != "0":
                return "£" + str(round_total)[:2] + "," + str(round_total)[2] + "bn"
            else:
                return "£" + str(round_total)[:2] + "bn"
        if total_len == 6:
            round_total = int(round(number, -3))
            if str(round_total)[3] != "0":
                return "£" + str(round_total)[:3] + "," + str(round_total)[3] + "bn"
            else:
                return "£" + str(round_total)[:3] + "bn"

        # this bit is for top250
        if total_len == 7:
            round_total = round(number, -5)
            return "£" + str(round_total)[:1] + "m"
        if total_len == 8:
            round_total = round(number, -6)
            return "£" + str(round_total)[:2] + "m"
        if total_len == 9:
            round_total = round(number, -7)
            return "£" + str(round_total)[:3] + "m"
        if total_len == 10:
            round_total = round(number, -8)
            if str(round_total)[1] != "0":
                return "£" + str(round_total)[:1] + "," + str(round_total)[1] + "bn"
            else:
                return "£" + str(round_total)[:1] + "bn"
        if total_len == 11:
            round_total = round(number, -8)
            if str(round_total)[2] != "0":
                return "£" + str(round_total)[:2] + "," + str(round_total)[2] + "bn"
            else:
                return "£" + str(round_total)[:2] + "bn"

    except ValueError:
        print("not number")


def cal_group_angle(dist_no: int, group: List[str], **kwargs):
    """helper function for dandelion data class.
    Calculates distribution of first circle around center."""
    g_ang = dist_no / len(group)  # group_ang and distribution number
    output_list = []
    for i in range(len(group)):
        output_list.append(g_ang * i)
    if "all" not in kwargs:
        del output_list[5]
    # del output_list[0]
    return output_list


#  switches the type of data displayed in dandelion graph output
#  change type to dandelion_type
def get_dandelion_type_total(master, kwargs) -> int or str:
    tp = kwargs["quarter"][0]  # only one quarter for dandelion
    if "type" in kwargs:
        if kwargs["type"] == "remaining":
            cost = CostData(master, **kwargs)  # group costs data
            cost.get_forecast_cost_profile()
            # return sum(cost.profiles[tp]["total"]) - sum(cost.profiles[tp]["std"])
            return sum(cost.profiles[tp]["total"])
        if kwargs["type"] == "spent":
            cost = CostData(master, **kwargs)
            cost.get_forecast_cost_profile()
            # return cost.c_totals[tp]['total'] - (sum(cost.profiles[tp]["total"]) - sum(cost.profiles[tp]["std"]))
            return cost.totals[tp]["total"] - sum(cost.profiles[tp]["total"])
        if kwargs["type"] == "income":
            cost = CostData(master, **kwargs)
            return cost.totals[tp]["income_total"]
        if kwargs["type"] == "benefits":
            benefits = BenefitsData(master, **kwargs)
            return benefits.b_totals[tp]["total"]
        if kwargs["type"] == "ps resource":
            resource = ResourceData(master, **kwargs)
            return resource.ps_resource
        if kwargs["type"] == "contract resource":
            resource = ResourceData(master, **kwargs)
            return resource.contractor_resource
        if kwargs["type"] == "total resource":
            resource = ResourceData(master, **kwargs)
            return resource.total_resource
        if kwargs["type"] == "funded resource":
            resource = ResourceData(master, **kwargs)
            return resource.funded

    else:
        cost = CostData(master, **kwargs)  # group costs data
        return cost.totals[tp]["total"]


def calculate_circle_edge(sro_rag, fwd_look):
    if fwd_look == "No Change Expected":
        return sro_rag
    if fwd_look == "Improving":
        now = RAG_RANKING_DICT_COLOUR[sro_rag]
        return RAG_RANKING_DICT_NUMBER[now + 1]
    if fwd_look == "Worsening":
        now = RAG_RANKING_DICT_COLOUR[sro_rag]
        return RAG_RANKING_DICT_NUMBER[now - 1]


class DandelionData:
    """
    Data class for dandelion info graphic. Output dictionary to d_data()
    """

    def __init__(self, master, **kwargs):
        self.master = master
        self.kwargs = kwargs
        self.group = []
        self.group_stage_switch = ""
        self.handle_group()
        self.quarter = self.kwargs['quarter'][0]
        self.d_data = {}
        self.get_data()

    def handle_group(self):
        if "group" in self.kwargs:
            group_stage_switch = "group"
        if "stage" in self.kwargs:
            group_stage_switch = "stage"

        self.group_stage_switch = group_stage_switch
        self.group = self.kwargs[group_stage_switch]

    def get_data(self):
        dca_confidence = STANDARDISE_DCA_KEYS[self.kwargs["report"]]

        if "angles" in self.kwargs:
            if len(self.kwargs["angles"]) == len(self.kwargs[self.group_stage_switch]):
                g_ang_l = self.kwargs["angles"]
            else:
                raise InputError(
                    "The number of groups and angles don't match. Stopping."
                )
        else:
            # Dandelion graph needs an algorithm to calculate the distribution
            # of group circles. The circles are placed and distributed left
            # to right around the center circle.
            g_ang_l = []
            # start_point needs to come down as numbers increase
            start_point = 290 * (
                (29 - ((len(self.kwargs[self.group_stage_switch])) - 2)) / 29
            )
            # distribution increase needs to come down as numbers increase
            distribution_start = 0
            distribution_increase = 140
            if (
                len(self.kwargs[self.group_stage_switch]) > 2
            ):  # no change in distribution increase if group of two
                for i in range(len(self.kwargs[self.group_stage_switch])):
                    distribution_increase = distribution_increase * 0.82
            for i in range(len(self.kwargs[self.group_stage_switch])):
                angle = distribution_start + start_point
                if angle > 360:
                    angle = angle - 360
                g_ang_l.append(int(angle))
                distribution_start += distribution_increase

        g_d = {}  # group dictionary. first outer circle.
        l_g_d = {}  # lower group dictionary

        pf_wlc = get_dandelion_type_total(self.master, self.kwargs)  # portfolio wlc
        if "pc" in self.kwargs:  # pc portfolio colour
            pf_colour = COLOUR_DICT[self.kwargs["pc"]]
            pf_colour_edge = COLOUR_DICT[self.kwargs["pc"]]
        else:
            pf_colour = "#FFFFFF"
            pf_colour_edge = "grey"
        pf_text = "Portfolio\n" + dandelion_number_text(pf_wlc, **self.kwargs)

        ## center circle
        g_d["portfolio"] = {
            "axis": (0, 0),
            "r": math.sqrt(pf_wlc),
            "colour": pf_colour,
            "text": pf_text,
            "fill": "solid",
            "ec": pf_colour_edge,
            "alignment": ("center", "center"),
        }

        ## first outer circle
        for i, g in enumerate(self.group):
            self.kwargs[self.group_stage_switch] = [g]
            if g == "pipeline":
                g_wlc = self.master.pipeline_dict["pipeline"]["wlc"]
            else:
                g_wlc = get_dandelion_type_total(self.master, self.kwargs)

            if "same_size" in self.kwargs:
                if self.kwargs["same_size"] == "Yes":
                    g_wlc = 46000

            g_text = (
                g + "\n" + dandelion_number_text(g_wlc, **self.kwargs)
            )  # group text

            if "values" in self.kwargs:
                if self.kwargs["values"] == "No":
                    g_text = g

            if len(self.group) > 1:
                y_axis = 0 + (
                    (math.sqrt(pf_wlc) * 3.25) * math.sin(math.radians(g_ang_l[i]))
                )
                x_axis = 0 + (math.sqrt(pf_wlc) * 2.75) * math.cos(
                    math.radians(g_ang_l[i])
                )

                if g_wlc < pf_wlc / 20:
                    g_wlc = pf_wlc / 20

                # c_colour = circle_colours[i]

                g_d[g] = {
                    "axis": (y_axis, x_axis),
                    "r": math.sqrt(g_wlc),
                    "wlc": g_wlc,
                    "colour": "#FFFFFF",
                    # "colour": c_colour,
                    "text": g_text,
                    "fill": "dashed",
                    "ec": "grey",
                    "alignment": ("center", "center"),
                    "angle": g_ang_l[i],
                }

            else:
                g_d = {}
                pf_wlc = g_wlc * 3
                # g_text = g + "\n" + dandelion_number_text(g_wlc)  # group text
                if g_wlc == 0:
                    g_wlc = 20
                g_d[g] = {
                    "axis": (0, 0),
                    "r": math.sqrt(g_wlc),
                    "wlc": g_wlc,
                    "colour": "#FFFFFF",
                    "text": g_text,
                    "fill": "dashed",
                    "ec": "grey",
                    "alignment": ("center", "center"),
                }

        if len(self.kwargs[self.group_stage_switch]) > 1:
            group_angles = []
            for g in self.kwargs[self.group_stage_switch]:
                group_angles.append((g, g_d[g]["angle"]))
            logger.info("The group circle angles are " + str(group_angles))

        ## second outer circle
        for i, g in enumerate(self.group):
            self.kwargs[self.group_stage_switch] = [g]
            if g == "pipeline":
                project_group = self.master.pipeline_list
            else:
                project_group = get_group(self.master, self.quarter, **self.kwargs)  # lower group
            p_list = []
            for p in project_group:
                self.kwargs[self.group_stage_switch] = [
                    p
                ]  # project level is in group not stage.
                if g == "pipeline":
                    p_value = self.master.pipeline_dict[p]["wlc"]
                else:
                    p_value = get_dandelion_type_total(
                        self.master, self.kwargs
                    )  # project wlc

                # if "report" not in self.kwargs:  # DON'T UNDERSTAND THIS IF STATEMENT
                if "something" in self.kwargs:  # SOMETHING TO DO WITH SCHEDULE?
                    if g == "pipeline":
                        p_schedule = self.master.pipeline_dict[p]["wlc"]
                    else:
                        quarter_index = get_quarter_index(self.master, self.quarter)
                        bc = BC_STAGE_DICT[
                            self.master.master_data[quarter_index]["data"][p][
                                "IPDC approval point"
                            ]
                        ]
                        ms = MilestoneData(self.master, **self.kwargs)
                        next_stage = NEXT_STAGE_DICT[bc]
                        try:
                            d = get_milestone_date(
                                self.master.abbreviations[p]["abb"],
                                ms.milestone_dict,
                                self.quarter,
                                next_stage,
                            )
                            p_schedule = (d - datetime.date.today()).days
                        except TypeError:
                            p_schedule = 10000000000
                            if "order_by" in self.kwargs:
                                if self.kwargs["order_by"] == "schedule":
                                    print("can't calculate " + p + "'s schedule")
                else:
                    p_schedule = 0
                p_list.append((p_value, p_schedule, p))
            if "order_by" in self.kwargs:
                if self.kwargs["order_by"] == "schedule":
                    p_list.sort(key=lambda x: x[1])
                    if g == "pipeline":
                        l_g_d[g] = list(reversed(p_list))
                    else:
                        l_g_d[g] = p_list
            else:
                l_g_d[g] = list(reversed(sorted(p_list)))

        for g in self.group:
            g_wlc = g_d[g]["wlc"]
            g_radius = g_d[g]["r"]
            g_y_axis = g_d[g]["axis"][0]  # group y axis
            g_x_axis = g_d[g]["axis"][1]  # group x axis
            try:
                p_values_list, p_schedule_list, p_list = zip(*l_g_d[g])
            except ValueError:  # handles no projects in l_g_d list
                continue
            if len(p_list) > 2:
                # if len(p_list) > 2 or len(self.kwargs[self.group_stage_switch]) == 1:
                ang_l = cal_group_angle(360, p_list, all=True)
            else:
                if len(p_list) == 1:
                    ang_l = [g_d[g]["angle"]]
                if len(p_list) == 2:
                    ang_l = [g_d[g]["angle"], g_d[g]["angle"] + 70]

            for i, p in enumerate(p_list):
                p_value = p_values_list[i]
                if "same_size" in self.kwargs:
                    if self.kwargs["same_size"] == "Yes":
                        p_value = 6000
                p_data = get_correct_p_data(self.master, p, self.quarter)
                try:  # this is for pipeline projects
                    if "confidence" in self.kwargs:  # change confidence type here
                        rag = p_data[DCA_KEYS[self.kwargs["confidence"]]]
                    else:
                        try:
                            rag = p_data[dca_confidence]
                        except KeyError:  # top35 has no rag data
                            rag = None
                    colour = COLOUR_DICT[convert_rag_text(rag)]  # bubble colour
                except TypeError:  # p_data is None for pipeline projects
                    colour = "#FFFFFF"
                if "circle_colour" in self.kwargs:
                    if self.kwargs["circle_colour"] == "No":
                        colour = FACE_COLOUR

                project_text = (
                    self.master['project_information'][p]["Abbreviations"]
                    + "\n"
                    + dandelion_number_text(p_value, **self.kwargs)
                )
                if "type" in self.kwargs:
                    if self.kwargs["type"] in [
                        "ps resource",
                        "contract resource",
                        "total resource",
                        "funded resource",
                    ]:
                        project_text = (
                            self.master.abbreviations[p]["abb"]
                            + ", "
                            + dandelion_number_text(p_value, **self.kwargs)
                        )
                if "values" in self.kwargs:
                    if self.kwargs["values"] == "No":
                        project_text = self.master.abbreviations[p]["abb"]
                if (
                    p_value < pf_wlc / 500
                ):  # achieve some consistency for zero / low values
                    p_value = pf_wlc / 500
                if colour == "#FFFFFF" or colour == FACE_COLOUR:
                    if p in self.master['meta_groupings'][self.quarter]['GMPP']:
                        edge_colour = "#000000"
                    else:
                        edge_colour = "grey"
                else:
                    if p in self.master['meta_groupings'][self.quarter]['GMPP']:
                        edge_colour = '#000000'
                        # edge_colour = COLOUR_DICT[p_data['SRO Forward Look Assessment']]
                    else:
                        edge_colour = colour
                        # edge_colour = COLOUR_DICT[p_data['SRO Forward Look Assessment']]
                if "circle_edge" in self.kwargs:
                    if self.kwargs["circle_edge"] == "forward_look":
                        try:
                            fwd_look = p_data["SRO Forward Look Assessment"]
                            edge_rag = calculate_circle_edge(rag, fwd_look)
                            edge_colour = COLOUR_DICT[edge_rag]
                        except KeyError:
                            raise InputError(
                                "No SRO Forward Look Assessment key in quarter master. "
                                "This key must be present for this dandelion command. Stopping."
                            )
                    if self.kwargs["circle_edge"] == "ipa":
                        edge_colour = COLOUR_DICT[p_data["GMPP - IPA DCA"]]

                try:
                    if len(p_list) >= 16:
                        multi = (pf_wlc / g_wlc) ** 1.1
                    elif 15 >= len(p_list) >= 11:
                        multi = (pf_wlc / g_wlc) ** (1.0 / 2.0)  # square root
                    #  only one/two bubbles don't distribute well
                    elif len(p_list) == 1 or len(p_list) == 2:
                        multi = 2.2
                    else:
                        if g_wlc / pf_wlc >= 0.33:
                            multi = (pf_wlc / g_wlc) ** (1.0 / 2)
                        else:
                            multi = (pf_wlc / g_wlc) ** (1.0 / 3.0)  # cube root

                    p_y_axis = g_y_axis + (g_radius * multi) * math.sin(
                        math.radians(ang_l[i])
                    )
                    p_x_axis = g_x_axis + (g_radius * multi) * math.cos(
                        math.radians(ang_l[i])
                    )
                except ZeroDivisionError:
                    p_y_axis = g_y_axis + 100 * math.sin(math.radians(ang_l[i]))
                    p_x_axis = g_x_axis + 100 * math.cos(math.radians(ang_l[i]))

                if 179 >= ang_l[i] >= 165:
                    text_angle = ("left", "top")
                if 195 >= ang_l[i] >= 181:
                    text_angle = ("right", "top")
                if ang_l[i] == 180:
                    text_angle = ("center", "top")
                if 5 >= ang_l[i] or 355 <= ang_l[i]:
                    text_angle = ("center", "bottom")
                if 164 >= ang_l[i] >= 6:
                    text_angle = ("left", "center")
                if 354 >= ang_l[i] >= 196:
                    text_angle = ("right", "center")

                try:
                    t_multi = (g_wlc / p_value) ** (1.0 / 4.0)
                    # print(p, t_multi)
                    if t_multi <= 1.3:
                        t_multi = 1.3
                except ZeroDivisionError:
                    t_multi = 1
                yx_text_position = (
                    p_y_axis
                    + (math.sqrt(p_value) * t_multi) * math.sin(math.radians(ang_l[i])),
                    p_x_axis
                    + (math.sqrt(p_value) * t_multi) * math.cos(math.radians(ang_l[i])),
                )

                g_d[p] = {
                    "axis": (p_y_axis, p_x_axis),
                    "r": math.sqrt(p_value),
                    "wlc": p_value,
                    "colour": colour,
                    "text": project_text,
                    "fill": "solid",
                    "ec": edge_colour,
                    "alignment": text_angle,
                    "tp": yx_text_position,
                }

        self.d_data = g_d


def dandelion_data_into_wb(d_data: DandelionData) -> workbook:
    """
    Simple function that returns data required for the dandelion graph.
    """
    wb = Workbook()
    for tp in d_data.d_data.keys():
        ws = wb.create_sheet(
            make_file_friendly(tp)
        )  # creating worksheets. names restricted to 30 characters.
        ws.title = make_file_friendly(tp)  # title of worksheet
        for i, project in enumerate(d_data.d_data[tp]["projects"]):
            ws.cell(row=2 + i, column=1).value = d_data.d_data[tp]["group"][i]
            ws.cell(row=2 + i, column=2).value = d_data.d_data[tp]["abb"][i]
            ws.cell(row=2 + i, column=3).value = project
            ws.cell(row=2 + i, column=4).value = int(d_data.d_data[tp]["cost"][i])
            ws.cell(row=2 + i, column=5).value = d_data.d_data[tp]["rag"][i]

        ws.cell(row=1, column=1).value = "Group"
        ws.cell(row=1, column=2).value = "Project"
        ws.cell(row=1, column=3).value = "Graph details"
        ws.cell(row=1, column=4).value = "WLC (forecast)"
        ws.cell(row=1, column=5).value = "DCA"

    wb.remove(wb["Sheet"])
    return wb


def make_a_dandelion_auto(dl: DandelionData, **kwargs):
    """function used to compile dandelion graph. Data is taken from
    DandelionData class."""

    fig, ax = plt.subplots(figsize=(10, 8), facecolor=FACE_COLOUR)

    if "circle_edge" in kwargs:
        if kwargs["circle_edge"] == "forward_look" or "ipa":
            linewidth = 2.0
    else:
        linewidth = 1.0

    for c in dl.d_data.keys():
        circle = plt.Circle(
            dl.d_data[c]["axis"],  # x, y position
            radius=dl.d_data[c]["r"],
            fc=dl.d_data[c]["colour"],  # face colour
            ec=dl.d_data[c]["ec"],  # edge colour
            linewidth=linewidth,
            zorder=2,
        )
        ax.add_patch(circle)
        try:
            ax.annotate(
                dl.d_data[c]["text"],  # text
                xy=dl.d_data[c]["axis"],  # x, y position
                xycoords="data",
                xytext=dl.d_data[c]["tp"],  # text position
                fontsize=7,
                fontname=FONT_TYPE,
                horizontalalignment=dl.d_data[c]["alignment"][0],
                verticalalignment=dl.d_data[c]["alignment"][1],
                zorder=3,
            )
        except KeyError:
            # key error will occur for first and second outer circles as different text
            ax.annotate(
                dl.d_data[c]["text"],  # text
                xy=dl.d_data[c]["axis"],  # x, y position
                fontsize=9,
                fontname=FONT_TYPE,
                horizontalalignment=dl.d_data[c]["alignment"][0],
                verticalalignment=dl.d_data[c]["alignment"][1],
                weight="bold",  # bold here as will be group text
                zorder=3,
            )

    # place lines
    line_clr = "#ececec"
    line_style = "dashed"
    for i, g in enumerate(dl.group):
        dl.kwargs[dl.group_stage_switch] = [g]
        ax.arrow(
            0,
            0,
            dl.d_data[g]["axis"][0],
            dl.d_data[g]["axis"][1],
            fc=line_clr,
            ec=line_clr,
            # linestyle=line_style,
            zorder=1,
        )

        if g == "pipeline":
            lower_g = dl.master.pipeline_list
        else:
            lower_g = get_group(dl.master, dl.kwargs['quarter'][0], **dl.kwargs)
        for p in lower_g:
            ax.arrow(
                dl.d_data[g]["axis"][0],
                dl.d_data[g]["axis"][1],
                dl.d_data[p]["axis"][0] - dl.d_data[g]["axis"][0],
                dl.d_data[p]["axis"][1] - dl.d_data[g]["axis"][1],
                fc=line_clr,
                ec=line_clr,
                # linestyle=line_style,
                zorder=1,
            )

    plt.axis("scaled")
    plt.axis("off")
    if kwargs["chart"] != "save":
        plt.show()

    return fig
