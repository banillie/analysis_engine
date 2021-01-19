# from project_analysis.project_reports import produce_word_doc
from datetime import datetime
import datetime

from other.oldegg_functions import baseline_index, baseline_information, \
    project_all_milestones_dict, get_ben_totals


def test_getting_milestone_baseline_data(diff_milestone_types):
    projects = diff_milestone_types[0].projects
    bl_info = baseline_information(projects, diff_milestone_types, 'Re-baseline IPDC milestones')
    bl_index = baseline_index(bl_info, diff_milestone_types)
    assert isinstance(bl_info, (dict,))
    assert bl_index["Apollo 11"] == [0, 1]


def test_get_milestone_data_with_diff_date_types(diff_milestone_types):
    projects = diff_milestone_types[0].projects
    bl_info = baseline_information(projects, diff_milestone_types, 'Re-baseline IPDC milestones')
    bl_index = baseline_index(bl_info, diff_milestone_types)
    milestones = project_all_milestones_dict(projects, diff_milestone_types, bl_index, 0)
    assert milestones['Apollo 13']['Gateway Review 5'] == {datetime.date(2027, 4, 26): 'tbc EIS'}


def test_get_benefits_baseline_data(benefits_master):
    projects = benefits_master[0].projects
    bl_info = baseline_information(projects, benefits_master, "Re-baseline IPDC benefits")
    bl_index = baseline_index(bl_info, benefits_master)
    assert isinstance(bl_info, (dict,))
    assert bl_index['Apollo 11'] == [0, 1, 0]


def test_get_benefits_data(benefits_master):
    projects = benefits_master[0].projects
    bl_info = baseline_information(projects, benefits_master, "Re-baseline IPDC benefits")
    bl_index = baseline_index(bl_info, benefits_master)
    benefits = get_ben_totals('Apollo 11', bl_index, benefits_master)
    assert type(benefits) == tuple
