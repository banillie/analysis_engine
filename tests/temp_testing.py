# from project_analysis.project_reports import produce_word_doc
from data_mgmt.data import Projects
from data_mgmt.oldegg_functions import baseline_index, baseline_information, \
    project_all_milestones_dict


def test_getting_baseline_data(diff_milestone_types):
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
    assert milestones == []


