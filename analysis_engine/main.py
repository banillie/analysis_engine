import argparse
from argparse import RawTextHelpFormatter
import itertools
import sys

from openpyxl import load_workbook

from analysis_engine.data import (
    get_master_data,
    Master,
    get_project_information,
    VfMData,
    root_path,
    vfm_into_excel,
    MilestoneData,
    put_milestones_into_wb,
    run_p_reports,
    RiskData,
    risks_into_excel,
    DcaData,
    dca_changes_into_excel,
    dca_changes_into_word,
    open_word_doc,
    Pickle,
    open_pickle_file,
    ipdc_dashboard,
    CostData,
    cost_v_schedule_chart_into_wb,
    make_file_friendly,
    DandelionData,
    dandelion_data_into_wb,
    put_matplotlib_fig_into_word,
    cost_profile_into_wb,
    cost_profile_graph,
    data_query_into_wb,
    get_data_query_key_names,
    ProjectNameError,
    milestone_chart,
    get_cost_stackplot_data,
    cost_stackplot_graph,
    cal_group,
    make_a_dandelion_auto, build_speedials,
)

import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s: %(levelname)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
)
logger = logging.getLogger(__name__)

DFT_GROUP = ["HSMRPG", "AMIS", "Rail", "RPE"]


def check_remove(sc_args):  # subcommand arg
    if sc_args["remove"]:
        from analysis_engine.data import CURRENT_LOG
        for p in sc_args["remove"]:
            if p + " successfully removed from analysis." not in CURRENT_LOG:
                logger.warning(p + ' not recognised and therefore not removed from analysis.'
                                   ' Please make sure "remove" entry is correct.')


def run_correct_args(
    m: Master,
    ae_class: CostData or VfMData or DcaData or RiskData,
    args: argparse.ArgumentParser,
) -> CostData or VfMData or DcaData or RiskData:
    if args["quarters"] and args["stage"] and args["remove"]:
        data = ae_class(
            m, quarter=args["quarters"], stage=args["stage"], remove=args["remove"]
        )
    elif args["quarters"] and args["group"] and args["remove"]:
        data = ae_class(
            m, quarter=args["quarters"], group=args["group"], remove=args["remove"]
        )
    elif args["quarters"] and args["stage"]:
        data = ae_class(m, quarter=args["quarters"], stage=args["stage"])
    elif args["quarters"] and args["group"]:
        data = ae_class(m, quarter=args["quarters"], group=args["group"])
    elif args["quarters"] and args["remove"]:
        data = ae_class(m, quarter=args["quarters"], remove=args["remove"])
    elif args["baselines"] and args["stage"] and args["remove"]:
        data = ae_class(
            m, baseline=args["baselines"], stage=args["stage"], remove=args["remove"]
        )
    elif args["baselines"] and args["group"] and args["remove"]:
        data = ae_class(
            m, baseline=args["baselines"], group=args["group"], remove=args["remove"]
        )
    elif args["baselines"] and args["stage"]:
        data = ae_class(m, baseline=args["baselines"], stage=args["stage"])
    elif args["baselines"] and args["group"]:
        data = ae_class(m, baseline=args["baselines"], group=args["group"])
    elif args["baselines"] and args["remove"]:
        data = ae_class(m, baseline=args["baselines"], remove=args["remove"])
    elif args["stage"] and args["remove"]:
        data = ae_class(
            m, quarter=["standard"], group=args["stage"], remove=args["remove"]
        )
    elif args["group"] and args["remove"]:
        data = ae_class(
            m, quarter=["standard"], group=args["group"], remove=args["remove"]
        )
    elif args["baselines"]:
        data = ae_class(m, baseline=args["baselines"])
    elif args["quarters"]:
        data = ae_class(m, quarter=args["quarters"])
    elif args["stage"]:
        data = ae_class(m, quarter=["standard"], group=args["stage"])
    elif args["group"]:
        data = ae_class(m, quarter=["standard"], group=args["group"])
    elif args["remove"]:  # HERE. NOT WORKING
        data = ae_class(m, quarter=["standard"], remove=args["remove"])
    else:
        data = ae_class(m, quarter=["standard"])

    return data


def get_args_for_file(args: argparse) -> list:
    l = []  # l is list
    for x in args.values():
        if x is not None:
            ffx = make_file_friendly(x)  # ffx
            l.append(ffx)
    l = l[1:-1]  # get rid of builtin_function_or_method
    unpack = itertools.chain.from_iterable(l)
    return list(unpack)


def initiate(args):
    print("creating a master data file for analysis_engine")
    try:
        master = Master(get_master_data(), get_project_information())
    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)

    path_str = str("{0}/core_data/pickle/master".format(root_path))
    Pickle(master, path_str)


def run_general(args):
    programme = args["subparser_name"]
    print("compiling " + programme + " analysis")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    try:
        if programme == "vfm":
            c = run_correct_args(m, VfMData, args)  # c is class
            wb = vfm_into_excel(c)
        if programme == "risks":
            c = run_correct_args(m, RiskData, args)
            wb = risks_into_excel(c)
        if programme == "dcas":
            c = run_correct_args(m, DcaData, args)
            wb = dca_changes_into_excel(c)
        if programme == "costs":
            doc = open_word_doc(root_path / "input/summary_temp.docx")
            c = run_correct_args(m, CostData, args)
            wb = cost_profile_into_wb(c)
            if args["chart"]:
                if args["title"]:
                    graph = cost_profile_graph(c, title=args["title"], chart=True)
                else:
                    graph = cost_profile_graph(c, chart=True)
                if args["chart"] == "save":
                    put_matplotlib_fig_into_word(doc, graph, size=6, transparent=False)
                    doc.save(root_path / "output/costs_chart.docx")
        if programme != "speedial":  # only excel outputs
            wb.save(root_path / "output/{}.xlsx".format(programme))
            print(programme + " analysis has been compiled. Enjoy!")
    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)

    # TODO optional_args produces a list of strings, each of which are to be in the output file name path.
    # optional_args = get_args_for_file(args)
    # wb.save(root_path / "output/{}_{}.xlsx".format(programme, optional_args))
    # print(programme + " analysis has been compiled. Enjoy!")


def speedials(args):
    print("compiling speed dial analysis. This one takes a little time.")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    ## quarters
    if args["quarters"] and args["stage"] and args["remove"]:
        data = DcaData(
            m, quarter=args["quarters"], stage=args["stage"], remove=args["remove"]
        )
    elif args["quarters"] and args["group"] and args["remove"]:
        data = DcaData(
            m, quarter=args["quarters"], group=args["group"], remove=args["remove"]
        )
    elif args["quarters"] and args["stage"]:
        data = DcaData(m, quarter=args["quarters"], stage=args["stage"])
    elif args["quarters"] and args["group"]:
        data = DcaData(m, quarter=args["quarters"], group=args["group"])
    elif args["quarters"] and args["remove"]:
        data = DcaData(m, quarter=args["quarters"], remove=args["remove"])

    ## baselines
    elif args["baselines"] and args["stage"] and args["remove"]:
        data = DcaData(
            m, baseline=args["baselines"], stage=args["stage"], remove=args["remove"]
        )
    elif args["baselines"] and args["group"] and args["remove"]:
        data = DcaData(
            m, baseline=args["baselines"], group=args["group"], remove=args["remove"]
        )
    elif args["baselines"] and args["stage"]:
        data = DcaData(m, baseline=args["baselines"], stage=args["stage"])
    elif args["baselines"] and args["group"]:
        data = DcaData(m, baseline=args["baselines"], group=args["group"])
    elif args["baselines"] and args["remove"]:
        data = DcaData(m, baseline=args["baselines"], remove=args["remove"])
    elif args["stage"] and args["remove"]:
        data = DcaData(
            m, quarter=["standard"], group=args["stage"], remove=args["remove"]
        )
    elif args["group"] and args["remove"]:
        data = DcaData(
            m, quarter=["standard"], group=args["group"], remove=args["remove"]
        )
    elif args["baselines"]:
        data = DcaData(m, baseline=args["baselines"])
    elif args["quarters"]:
        data = DcaData(m, quarter=args["quarters"])
    elif args["stage"]:
        data = DcaData(m, quarter=["standard"], group=args["stage"])
    elif args["group"]:
        data = DcaData(m, quarter=["standard"], group=args["group"])
    elif args["remove"]:  # HERE. NOT WORKING
        data = DcaData(m, quarter=["standard"], remove=args["remove"])
    else:
        data = DcaData(m, quarter=["standard"])

    data.get_changes()
    hz_doc = open_word_doc(root_path / "input/summary_temp.docx")
    doc = dca_changes_into_word(data, hz_doc)
    doc.save(root_path / "output/speed_dials.docx")
    land_doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
    build_speedials(data, land_doc)
    land_doc.save(root_path / "output/speedial_graph.docx")
    print("Speed dial analysis has been compiled. Enjoy!")


def milestones(args):
    # args available for milestones quarters, baseline, stage, group, remove, type, dates, koi
    print("compiling milestone analysis_engine")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    try:
        if args["baselines"] and args["stage"] and args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baselines"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args["baselines"] and args["group"] and args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baselines"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args["quarters"] and args["group"] and args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args["quarters"] and args["stage"] and args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])

        elif args["stage"] and args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args["group"] and args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args['baseline'] and args["stage"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["stage"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baseline"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["stage"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["stage"] and args["dates"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], dates=args["dates"])
        elif args['baseline'] and args["group"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["group"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["group"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["group"] and args["dates"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], dates=args["dates"])
        elif args['quarters'] and args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args['quarters'] and args["stage"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args['quarters'] and args["type"] and args["stage"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args['quarters'] and args["type"] and args["remove"] and args["stage"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args['quarters'] and args["type"] and args["remove"] and args["stage"] and args["dates"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], dates=args["dates"])
        elif args['quarters'] and args["group"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args['quarters'] and args["type"] and args["group"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args['quarters'] and args["type"] and args["remove"] and args["group"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args['quarters'] and args["type"] and args["remove"] and args["group"] and args["dates"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], dates=args["dates"])

        elif args["type"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args["stage"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args["type"] and args["stage"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], quarter=["standard"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args["type"] and args["remove"] and args["stage"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args["type"] and args["remove"] and args["stage"] and args["dates"]:
            ms = MilestoneData(m, group=args["stage"], quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], dates=args["dates"])
        elif args["group"] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args["type"] and args["group"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], quarter=["standard"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args["type"] and args["remove"] and args["group"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args["type"] and args["remove"] and args["group"] and args["dates"]:
            ms = MilestoneData(m, group=args["group"], quarter=["standard"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], dates=args["dates"])

        elif args['baseline'] and args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])

        # elif args['baseline'] and args["stage"] and args["remove"] and args["dates"] and args["koi"]:
        #     ms = MilestoneData(m, group=args["stage"], baseline=args["baseline"], remove=args["remove"])
        #     ms.filter_chart_info(dates=args["dates"], key=args["koi"])

        elif args['baseline'] and args["type"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=DFT_GROUP, baseline=args["baseline"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["koi"]:
            ms = MilestoneData(m, group=DFT_GROUP, baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["dates"]:
            ms = MilestoneData(m, group=DFT_GROUP, baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"], dates=args["dates"])
        elif args['baseline'] and args["stage"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baseline"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args['baseline'] and args["stage"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["stage"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"])

        elif args['baseline'] and args["type"] and args["stage"] and args["koi"]:
            ms = MilestoneData(m, baseline=args["baseline"], group=DFT_GROUP)
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args["baselines"] and args["stage"] and args["dates"] and args["type"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"])
        elif args['baseline'] and args["stage"] and args["remove"] and args["type"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"])
        elif args['baseline'] and args["group"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args['baseline'] and args["remove"] and args["group"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(key=args["koi"])
        elif args['baseline'] and args["remove"] and args["group"] and args["dates"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"])
        elif args['baseline'] and args["type"] and args["group"] and args["koi"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"])
            ms.filter_chart_info(type=args["type"], key=args["koi"])
        elif args['baseline'] and args["type"] and args["remove"] and args["group"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baseline"], remove=args["remove"])
            ms.filter_chart_info(type=args["type"])


        # here



        elif args["baselines"] and args["group"] and args["dates"] and args["type"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"])

        elif args["baselines"] and args["stage"] and args["dates"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"])

        elif args["baselines"] and args["group"] and args["dates"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"])

        elif args["baselines"] and args["dates"] and args["type"]:
            ms = MilestoneData(m, baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"])

        elif args["baselines"] and args["group"] and args["type"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baselines"])
            ms.filter_chart_info(type=args["type"])

        elif args["baselines"] and args["stage"] and args["type"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baselines"])
            ms.filter_chart_info(type=args["type"])


        ## AND HERE
        elif args["remove"] and args["dates"] and args["koi"]:
            ms = MilestoneData(m, quarter=[str(m.current_quarter)], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"], key=args["koi"])
        elif args["remove"] and args["dates"]:
            ms = MilestoneData(m, quarter=[str(m.current_quarter)], remove=args["remove"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["koi"]:
            ms = MilestoneData(m, quarter=[str(m.current_quarter)])
            ms.filter_chart_info(key=args["koi"])



        elif args["baselines"] and args["dates"]:
            ms = MilestoneData(m, baseline=args["baselines"])
            ms.filter_chart_info(dates=args["dates"])

        elif args["baselines"] and args["type"]:
            ms = MilestoneData(m, baseline=args["baselines"])
            ms.filter_chart_info(type=args["type"])

        elif args["baselines"] and args["group"]:
            ms = MilestoneData(m, group=args["group"], baseline=args["baselines"])

        elif args["baselines"] and args["stage"]:
            ms = MilestoneData(m, group=args["stage"], baseline=args["baselines"])

        elif args["baselines"]:
            ms = MilestoneData(m, baseline=args["baselines"])

        # quarters
        elif args["quarters"] and args["stage"] and args["dates"] and args["type"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"])

        elif args["quarters"] and args["group"] and args["dates"] and args["type"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"])

        elif args["quarters"] and args["stage"] and args["dates"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"])

        elif args["quarters"] and args["stage"] and args["type"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"])
            ms.filter_chart_info(type=args["type"])

        elif args["quarters"] and args["group"] and args["dates"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"])

        elif args["quarters"] and args["group"] and args["type"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"])
            ms.filter_chart_info(type=args["type"])

        elif args["quarters"] and args["dates"] and args["type"]:
            ms = MilestoneData(m, quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"])

        elif args["quarters"] and args["type"]:
            ms = MilestoneData(m, quarter=args["quarters"])
            ms.filter_chart_info(dates=args["type"])

        elif args["quarters"] and args["dates"]:
            ms = MilestoneData(m, quarter=args["quarters"])
            ms.filter_chart_info(dates=args["dates"])

        elif args["quarters"] and args["group"]:
            ms = MilestoneData(m, group=args["group"], quarter=args["quarters"])

        elif args["quarters"] and args["stage"]:
            ms = MilestoneData(m, group=args["stage"], quarter=args["quarters"])

        elif args["quarters"]:
            ms = MilestoneData(m, quarter=args["quarters"])

        # dates
        elif args["dates"] and args["group"]:
            ms = MilestoneData(m, quarter=["standard"], group=args["group"])
            ms.filter_chart_info(dates=args["dates"])

        elif args["dates"] and args["stage"]:
            ms = MilestoneData(m, quarter=["standard"], group=args["stage"])
            ms.filter_chart_info(dates=args["dates"])

        elif args["dates"] and args["type"]:
            ms = MilestoneData(m, quarter=["standard"])
            ms.filter_chart_info(dates=args["dates"], type=args["type"])

        # type
        elif args["type"] and args["stage"]:
            ms = MilestoneData(m, quarter=["standard"], group=args["stage"])
            ms.filter_chart_info(type=args["type"])

        elif args["type"] and args["group"]:
            ms = MilestoneData(m, quarter=["standard"], group=args["group"])
            ms.filter_chart_info(type=args["type"])

        elif args["type"]:
            ms = MilestoneData(m, quarter=["standard"])
            ms.filter_chart_info(type=args["type"])
        elif args["dates"]:
            ms = MilestoneData(m, quarter=["standard"])
            ms.filter_chart_info(dates=args["dates"])
        elif args["group"]:
            ms = MilestoneData(m, quarter=["standard"], group=args["group"])
        elif args["stage"]:
            ms = MilestoneData(m, quarter=["standard"], group=args["stage"])

        else:
            ms = MilestoneData(m, quarter=["standard"])

        wb = put_milestones_into_wb(ms)
        wb.save(root_path / "output/milestone_data_output.xlsx")

        if args["chart"]:
            if args["title"] and args["blue_line"]:
                if args["blue_line"] == "today":
                    graph = milestone_chart(
                        ms, title=args["title"], blue_line="Today", chart=True
                    )
                elif args["blue_line"] == "ipdc":
                    graph = milestone_chart(
                        ms, title=args["title"], blue_line="ipdc_date", chart=True
                    )
                else:
                    graph = milestone_chart(
                        ms, title=args["title"], blue_line=args["blue_line"], chart=True
                    )
            elif args["title"]:
                graph = milestone_chart(ms, title=args["title"], chart=True)
            elif args["blue_line"]:
                graph = milestone_chart(ms, blue_line="Today", chart=True)
            else:
                graph = milestone_chart(ms, chart=True)
            if args["chart"] == "save":
                doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
                put_matplotlib_fig_into_word(doc, graph, size=8, transparent=False)
                doc.save(root_path / "output/milestones_chart.docx")

    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)
    except Warning:
        logger.critical("To many milestone for chart. Stopping")
        sys.exit(1)


def summaries(args):
    print("compiling summaries")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    if args["group"]:
        run_p_reports(m, group=args["group"], baseline=["standard"])
    else:
        run_p_reports(m, baseline=["standard"])


def dashboard(args):
    print("compiling ipdc dashboards")
    dashboard_master = load_workbook(root_path / "input/dashboards_master.xlsx")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    wb = ipdc_dashboard(m, dashboard_master)
    wb.save(root_path / "output/completed_ipdc_dashboard.xlsx")
    print("dashboard compiled. enjoy!")


def matrix(args):
    print("compiling cost and schedule matrix analysis")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    costs = CostData(m, m.current_projects)
    miles = MilestoneData(m, m.current_projects)
    miles.calculate_schedule_changes()
    wb = cost_v_schedule_chart_into_wb(miles, costs)
    wb.save(root_path / "output/costs_schedule_matrix.xlsx")
    print("Cost and schedule matrix compiled. Enjoy!")


def costs_sp(args):
    try:
        print("compiling cost stackplot")
        m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
        if args["group"]:
            g = cal_group(args["group"], m, 0)
            sp_data = get_cost_stackplot_data(m, g, str(m.current_quarter), type="comp")
        # elif args["stage"]:
        #     s = cal_group(args["stage"], m, 0)
        #     sp_data = get_cost_stackplot_data(m, s, str(m.current_quarter), type="comp")
        else:
            sp_data = get_cost_stackplot_data(
                m,
                ["HSMRPG", "Rail", "RPE", "AMIS"],
                str(m.current_quarter),
                type="comp",
            )
        sp_graph = cost_stackplot_graph(sp_data)
        doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
        put_matplotlib_fig_into_word(doc, sp_graph, size=7.5)
        doc.save(root_path / "output/stackplot_graph.docx")
    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)


# note query option needs to work for one quarter only.
def query(args):
    print("Getting data")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    if args["keys"]:
        wb = data_query_into_wb(m, keys=args["keys"], quarter=["standard"])
        wb.save(root_path / "output/data_query.xlsx")
        print("Data compiled. Enjoy!")
    elif args["file_name"]:
        l = get_data_query_key_names(
            root_path / "input/{}.csv".format(args["file_name"])
        )
        wb = data_query_into_wb(m, keys=l, quarter=["standard"])
        wb.save(root_path / "output/data_query.xlsx")
        print("Data compiled using " + args["file_name"] + ".cvs file. Enjoy!")
    elif args["file_name"] and args["quarters"]:
        l = get_data_query_key_names(
            root_path / "input/{}.csv".format(args["file_name"])
        )
        wb = data_query_into_wb(m, keys=l, quarter=args["quarters"])
        wb.save(root_path / "output/data_query.xlsx")
        print("Data compiled using " + args["file_name"] + ".cvs file. Enjoy!")
    else:
        l = get_data_query_key_names(root_path / "input/key_names.csv")
        wb = data_query_into_wb(m, keys=l)
        wb.save(root_path / "output/data_query.xlsx")
        print("Data compiled using key_names cvs file. Enjoy!")


def dandelion(args):
    # args available for dandelion "quarter", "baseline", "remove", "type", "stage"
    print("compiling dandelion analysis")
    m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
    try:
        # if args['baseline'] and args['stage'] and args['type'] and args["remove"]:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], group=args["stage"], meta=args["type"], remove=args["remove"]
        #     )
        # elif args['baseline'] and args['group'] and args['type'] and args["remove"]:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], group=args["group"], meta=args["type"], remove=args["remove"]
        #     )
        if args['quarters'] and args['stage'] and args['type'] and args["remove"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], group=args["stage"], meta=args["type"], remove=args["remove"]
            )
        elif args['quarters'] and args['group'] and args['type'] and args["remove"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], group=args["group"], meta=args["type"], remove=args["remove"]
            )
        elif args['stage'] and args["type"] and args['remove']:
            d_data = DandelionData(
                m, meta=args["type"], group=args["stage"], remove=args["remove"]
            )
        elif args['group'] and args["type"] and args['remove']:
            d_data = DandelionData(
                m, remove=args["remove"], group=args["group"], meta=args["type"]
            )
        # elif args['baseline'] and args['type'] and args['remove']:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], remove=args["remove"], meta=args["type"]
        #     )
        # elif args['baseline'] and args['stage'] and args['type']:
        #     d_data = DandelionData(
        #         m, group=args["stage"], meta=args["type"], baseline=args["baseline"]
        #     )
        # elif args['baseline'] and args['stage'] and args['remove']:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], group=args["stage"], remove=args["remove"]
        #     )
        # elif args['baseline'] and args["group"] and args["remove"]:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], group=args["group"], remove=args["remove"]
        #     )
        elif args["quarters"] and args["stage"] and args["remove"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], group=args["stage"], remove=args["remove"]
            )
        elif args["quarters"] and args["type"] and args["remove"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], meta=args["type"], remove=args["remove"]
            )
        elif args["quarters"] and args["group"] and args["remove"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], group=args["group"], remove=args["remove"]
            )
        elif args["quarters"] and args["stage"] and args["type"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], group=args["stage"], meta=args["type"]
            )
        elif args["quarters"] and args["group"] and args["type"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], group=args["group"], meta=args["type"]
            )
        elif args["quarters"] and args["group"]:
            d_data = DandelionData(m, quarter=args["quarters"], group=args["group"])
        elif args["quarters"] and args["stage"]:
            d_data = DandelionData(m, quarter=args["quarters"], group=args["stage"])
        elif args["quarters"] and args["type"]:
            d_data = DandelionData(m, quarter=args["quarters"], meta=args["type"])
        elif args["quarters"] and args["remove"]:
            d_data = DandelionData(m, quarter=args["quarters"], group=DFT_GROUP, remove=args["remove"])
        # elif args["baseline"] and args["stage"]:
        #     d_data = DandelionData(m, baseline=args["baseline"], group=args["stage"])
        # elif args["baseline"] and args["group"]:
        #     d_data = DandelionData(m, baseline=args["baseline"], group=args["group"])
        # elif args["baseline"] and args["type"]:
        #     d_data = DandelionData(m, baseline=args["baseline"], meta=args["type"])
        # elif args["baseline"] and args["remove"]:
        #     d_data = DandelionData(m, baseline=args["baseline"], group=DFT_GROUP, remove=args["remove"])
        elif args["group"] and args["type"]:
            d_data = DandelionData(
                m, quarter=[str(m.current_quarter)], group=args["group"], meta=args["type"]
            )
        elif args["stage"] and args["type"]:
            d_data = DandelionData(
                m, quarter=[str(m.current_quarter)], group=args["stage"], meta=args["type"]
            )
        elif args["stage"] and args["remove"]:
            d_data = DandelionData(
                m, quarter=[str(m.current_quarter)], group=args["stage"], remove=args["remove"]
            )
        elif args["group"] and args["remove"]:
            d_data = DandelionData(
                m, quarter=[str(m.current_quarter)], group=args["group"], remove=args["remove"]
            )
        elif args["type"] and args["remove"]:
            d_data = DandelionData(
                m, quarter=[str(m.current_quarter)], meta=args["type"], remove=args["remove"]
            )
        # elif args["baseline"] and args["remove"]:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], remove=args["remove"]
        #     )
        # elif args["baseline"] and args["type"]:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], meta=args["type"]
        #     )
        # elif args["baseline"] and args["stage"]:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], group=args["stage"]
        #     )
        # elif args["baseline"] and args["group"]:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], groupo=args["group"]
        #     )
        elif args["quarters"] and args["remove"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], remove=args["remove"]
            )
        # elif args["baseline"] and args["type"]:
        #     d_data = DandelionData(
        #         m, quarter=args["quarters"], meta=args["type"]
        #     )
        # elif args["baseline"] and args["stage"]:
        #     d_data = DandelionData(
        #         m, quarter=args["quarters"], group=args["stage"]
        #     )
        # elif args["baseline"] and args["group"]:
        #     d_data = DandelionData(
        #         m, quarter=args["quarters"], groupo=args["group"]
        #     )
        # elif args["baseline"]:
        #     d_data = DandelionData(
        #         m, baseline=args["baseline"], group=DFT_GROUP)
        elif args["quarters"]:
            d_data = DandelionData(
                m, quarter=args["quarters"], group=DFT_GROUP)
        elif args["group"]:
            d_data = DandelionData(
                m, quarter=[str(m.current_quarter)], group=args["group"]
            )
        elif args["stage"]:
            d_data = DandelionData(
                m, quarter=[str(m.current_quarter)], group=args["stage"]
            )
        elif args["type"]:
            d_data = DandelionData(
                m,
                quarter=[str(m.current_quarter)],
                group=DFT_GROUP
                ,
                meta=args["type"],
            )
        elif args["remove"]:
            d_data = DandelionData(
                m,
                quarter=[str(m.current_quarter)],
                group=DFT_GROUP
                ,
                remove=args["remove"],
            )
        else:
            d_data = DandelionData(
                m,
                quarter=[str(m.current_quarter)],
                group=DFT_GROUP
                ,
            )

        if args["chart"]:
            # if args["title"]:
            #     graph = make_a_dandelion_auto(d_data, title=args["title"], chart=True)
            # else:
            graph = make_a_dandelion_auto(d_data, chart=True)
            if args["chart"] == "save":
                doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
                put_matplotlib_fig_into_word(doc, graph, size=7.5)
                doc.save(root_path / "output/dandelion_graph.docx")
                print("Dandelion chart has been compiled")

        check_remove(args)

    except ProjectNameError as e:
        logger.critical(e)
        sys.exit(1)


def main():
    ae_description = "Welcome to the DfT Major Projects Portfolio Office analysis engine.\n\n" \
                     "To operate use subcommands outlined below. To navigate each subcommand\n" \
                     "option use the --help flag which will provide instructions on which optional\n" \
                     "arguments can be used with each subcommand. e.g. analysis dandelion --help."
    parser = argparse.ArgumentParser(
        prog="engine", description=ae_description, formatter_class=RawTextHelpFormatter
    )
    subparsers = parser.add_subparsers(dest="subparser_name")
    subparsers.metavar = "subcommand                "
    parser_initiate = subparsers.add_parser(
        "initiate", help="creates a master data file"
    )
    dashboard_description = "Creates IPDC dashboards. There are no optional arguments for this command.\n\n" \
                            "A blank master dashboard titled dashboards_master.xlsx must be in input file.\n\n" \
                            "A completed dashboard title completed_ipdc_dashboard.xlsx will be placed into\n" \
                            "the output file."
    parser_dashboard = subparsers.add_parser(
        "dashboards", description=dashboard_description, formatter_class=RawTextHelpFormatter
    )
    parser_dandelion = subparsers.add_parser(
        "dandelion",
        help="dandelion graph and data (early version of graph output).",
    )
    parser_costs = subparsers.add_parser(
        "costs",
        help="cost trend profile graph and data.",
    )
    parser_costs_sp = subparsers.add_parser(
        "costs_sp",
        help="cost stackplot graph and data (early version needs more testing).",
    )
    parser_milestones = subparsers.add_parser(
        "milestones",
        help="milestone schedule graphs and data.",
    )
    parser_vfm = subparsers.add_parser("vfm", help="vfm analysis")
    parser_summaries = subparsers.add_parser("summaries", help="summary reports")
    parser_risks = subparsers.add_parser("risks", help="risk analysis")
    parser_dca = subparsers.add_parser("dcas", help="dca analysis")
    parser_speedial = subparsers.add_parser("speedial", help="speed dial analysis")
    parser_matrix = subparsers.add_parser("matrix", help="cost v schedule chart")
    parser_data_query = subparsers.add_parser(
        "query", help="return data from core data"
    )

    # Arguments
    # stage
    for sub in [
        parser_dca,
        parser_vfm,
        parser_risks,
        parser_speedial,
        parser_dandelion,
        parser_costs,
        parser_costs_sp,
        # parser_data_query,
        parser_milestones,
    ]:
        sub.add_argument(
            "--stage",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
            help="Returns analysis for only those projects at the specified planning stage(s). User must enter one "
            'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
        )
    # group
    for sub in [
        parser_dca,
        parser_vfm,
        parser_risks,
        parser_speedial,
        parser_dandelion,
        parser_costs,
        # parser_data_query,
        parser_milestones,
        parser_summaries,
        parser_costs_sp,
    ]:
        sub.add_argument(
            "--group",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            help="Returns analysis for specified project(s), only. User must enter one or a combination of "
            'DfT Group names; "HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym or full name.',
        )
    # remove
    for sub in [
        parser_dca,
        parser_vfm,
        parser_risks,
        parser_speedial,
        parser_dandelion,
        parser_costs,
        # parser_data_query,
        parser_milestones,
    ]:
        sub.add_argument(
            "--remove",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            help="Removes specified projects from analysis. User must enter one or a combination of either"
            " a recognised DfT Group name, a recognised planning stage or the project(s) acronym or full"
            " name.",
        )
    # quarter
    for sub in [
        parser_dca,
        parser_vfm,
        parser_risks,
        parser_speedial,
        parser_dandelion,
        parser_costs,
        parser_data_query,
        parser_milestones,
    ]:
        sub.add_argument(
            "--quarters",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            help='Returns analysis for specified quarters. User must use correct format e.g "Q3 19/20"',
        )
    # baseline
    for sub in [
        parser_dca,
        parser_vfm,
        parser_risks,
        parser_speedial,
        # parser_dandelion,
        parser_costs,
        # parser_data_query,
        parser_milestones,
    ]:
        sub.add_argument(
            "--baselines",
            type=str,
            metavar="",
            action="store",
            nargs="+",
            choices=[
                "current",
                "last",
                "bl_one",
                "bl_two",
                "bl_three",
                "standard",
                "all",
            ],
            help="Returns analysis for specified baselines. User must use correct format"
            ' which are "current", "last", "bl_one", "bl_two", "bl_three", "standard", "all".'
            ' The "all" option returns all, "standard" returns first three',
        )

    parser_milestones.add_argument(
        "--type",
        type=str,
        metavar="",
        action="store",
        nargs="+",
        choices=["Approval", "Assurance", "Delivery"],
        help="Returns analysis for specified type of milestones.",
    )

    parser_milestones.add_argument(
        "--koi",
        type=str,
        # metavar="Key Name",
        action="store",
        nargs="+",
        help="Returns the specified keys of interest (KOI).",
    )

    parser_milestones.add_argument(
        "--dates",
        type=str,
        metavar="",
        action="store",
        nargs=2,
        help="dates for analysis. Must provide start date and then end date in format e.g. '1/1/2021' '1/1/2022'.",
    )

    parser_dandelion.add_argument(
        "--type",
        type=str,
        metavar="",
        action="store",
        choices=["spent", "remaining", "benefits"],
        help="Provide the type of value to include in dandelion e.g spent, remaining",
    )

    # chart
    for sub in [parser_dandelion, parser_costs, parser_milestones]:
        sub.add_argument(
            "--chart",
            type=str,
            metavar="",
            action="store",
            choices=["show", "save"],
            help="options for building and saving graph output. Commands are 'show' or 'save' ",
        )

    # title
    for sub in [parser_costs, parser_milestones]:
        sub.add_argument(
            "--title",
            type=str,
            metavar="",
            action="store",
            help="provide a title for chart. Optional",
        )

    parser_milestones.add_argument(
        "--blue_line",
        type=str,
        metavar="",
        action="store",
        help='Insert blue line into chart to represent a date. Options are "today" "ipdc" or a date in correct format e.g. "1/1/2021".',
    )

    parser_data_query.add_argument(
        "--keys",
        type=str,
        metavar="Key Name",
        action="store",
        nargs="+",
        help="Returns the specified data keys.",
    )

    parser_data_query.add_argument(
        "--file_name",
        type=str,
        action="store",
        help="provide name of csc file contain key names",
    )

    # for sub in [
    #     parser_dca,
    #     parser_vfm,
    #     parser_risks,
    #     parser_speedial,
    #     parser_dandelion,
    #     parser_costs,
    #     parser_data_query,
    #     parser_milestones,
    # ]:
    #     # all sub-commands have the same optional args. This is working
    #     # but prob could be refactored.
    #     sub.add_argument(
    #         "--stage",
    #         type=str,
    #         metavar="",
    #         action="store",
    #         nargs="+",
    #         choices=["FBC", "OBC", "SOBC", "pre-SOBC"],
    #         help="Returns analysis for those projects at the specified planning stage(s). Must be one "
    #         'or combination of "FBC", "OBC", "SOBC", "pre-SOBC".',
    #     )
    #     sub.add_argument(
    #         "--group",
    #         type=str,
    #         metavar="",
    #         action="store",
    #         nargs="+",
    #         # choices=DFT_GROUP
    #         ,
    #         help="Returns summaries for specified project(s). User can either input DfT Group name; "
    #         '"HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym',
    #     )
    #     # no quarters in dandelion yet
    #     sub.add_argument(
    #         "--quarters",
    #         type=str,
    #         metavar="",
    #         action="store",
    #         nargs="+",
    #         help="Returns analysis for specified quarters. Must be in format e.g Q3 19/20",
    #     )
    #     sub.add_argument(
    #         "--remove",
    #         type=str,
    #         metavar="",
    #         action="store",
    #         nargs="+",
    #         # choices=DFT_GROUP
    #         ,
    #         help="Removes specified projects from analysis. User can either input DfT Group name; "
    #              '"HSMRPG", "AMIS", "Rail", "RPE", or the project(s) acronym"',
    #     )

    parser_initiate.set_defaults(func=initiate)
    parser_dashboard.set_defaults(func=dashboard)
    parser_dandelion.set_defaults(func=dandelion)
    parser_costs.set_defaults(func=run_general)
    parser_vfm.set_defaults(func=run_general)
    parser_milestones.set_defaults(func=milestones)
    parser_summaries.set_defaults(func=summaries)
    parser_risks.set_defaults(func=run_general)
    parser_dca.set_defaults(func=run_general)
    parser_speedial.set_defaults(func=speedials)
    parser_matrix.set_defaults(func=matrix)
    parser_data_query.set_defaults(func=query)
    parser_costs_sp.set_defaults(func=costs_sp)
    args = parser.parse_args()
    # print(vars(args))
    args.func(vars(args))


if __name__ == "__main__":
    main()

('quarter',)
('baseline',)
('group',)
('stage',)
('type',)
('remove',)
('dates',)
('koi',)
('quarter', 'group')
('quarter', 'stage')
('quarter', 'type')
('quarter', 'remove')
('quarter', 'dates')
('quarter', 'koi')
('baseline', 'group')
('baseline', 'stage')
('baseline', 'type')
('baseline', 'remove')
('baseline', 'dates')
('baseline', 'koi')
('group', 'type')
('group', 'remove')
('group', 'dates')
('group', 'koi')
('stage', 'type')
('stage', 'remove')
('stage', 'dates')
('stage', 'koi')
('type', 'remove')
('type', 'dates')
('type', 'koi')
('remove', 'dates')
('remove', 'koi')
('dates', 'koi')
('quarter', 'group', 'stage')
('quarter', 'group', 'type')
('quarter', 'group', 'remove')
('quarter', 'group', 'dates')
('quarter', 'group', 'koi')
('quarter', 'stage', 'type')
('quarter', 'stage', 'remove')
('quarter', 'stage', 'dates')
('quarter', 'stage', 'koi')
('quarter', 'type', 'remove')
('quarter', 'type', 'dates')
('quarter', 'type', 'koi')
('quarter', 'remove', 'dates')
('quarter', 'remove', 'koi')
('quarter', 'dates', 'koi')
('baseline', 'group', 'stage')
('baseline', 'group', 'type')
('baseline', 'group', 'remove')
('baseline', 'group', 'dates')
('baseline', 'group', 'koi')
('baseline', 'stage', 'type')
('baseline', 'stage', 'remove')
('baseline', 'stage', 'dates')
('baseline', 'stage', 'koi')
('baseline', 'type', 'remove')
('baseline', 'type', 'dates')
('baseline', 'type', 'koi')
('baseline', 'remove', 'dates')
('baseline', 'remove', 'koi')
('baseline', 'dates', 'koi')
('group', 'type', 'remove')
('group', 'type', 'dates')
('group', 'type', 'koi')
('group', 'remove', 'dates')
('group', 'remove', 'koi')
('group', 'dates', 'koi')
('stage', 'type', 'remove')
('stage', 'type', 'dates')
('stage', 'type', 'koi')
('stage', 'remove', 'dates')
('stage', 'remove', 'koi')
('stage', 'dates', 'koi')
('type', 'remove', 'dates')
('type', 'remove', 'koi')
('type', 'dates', 'koi')
('remove', 'dates', 'koi')
('quarter', 'group', 'type', 'remove')
('quarter', 'group', 'type', 'dates')
('quarter', 'group', 'type', 'koi')
('quarter', 'group', 'remove', 'dates')
('quarter', 'group', 'remove', 'koi')
('quarter', 'group', 'dates', 'koi')
('quarter', 'stage', 'type', 'remove')
('quarter', 'stage', 'type', 'dates')
('quarter', 'stage', 'type', 'koi')
('quarter', 'stage', 'remove', 'dates')
('quarter', 'stage', 'remove', 'koi')
('quarter', 'stage', 'dates', 'koi')
('quarter', 'type', 'remove', 'dates')
('quarter', 'type', 'remove', 'koi')
('quarter', 'type', 'dates', 'koi')
('quarter', 'remove', 'dates', 'koi')














