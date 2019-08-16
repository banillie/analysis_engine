import pytest
import logging

from ..core import Master, Quarter, ProjectData


def test_master(master):
    q1_2017 = Quarter(1, 2017)
    m = Master(q1_2017, master)
    assert m.path == '/tmp/bcompiler/output/master.xlsx'
    assert m.filename == 'master.xlsx'
    assert str(m.quarter) == 'Q1 17/18'
    assert m.year == 2017
    assert len(m.projects) == 3

    assert m.projects[0] == 'PROJECT/PROGRAMME NAME 1'
    assert m.projects[1] == 'PROJECT/PROGRAMME NAME 2'
    assert m.projects[2] == 'PROJECT/PROGRAMME NAME 3'

    assert len(m['PROJECT/PROGRAMME NAME 1']) == 1276
    assert m['PROJECT/PROGRAMME NAME 1']['SRO Full Name'] == 'SRO FULL NAME 1'

    p1 = m['PROJECT/PROGRAMME NAME 1']
    assert p1['SRO Full Name'] == 'SRO FULL NAME 1'


def test_project_data_object(master):
    q2_2018 = Quarter(2, 2018)
    m = Master(q2_2018, master)
    assert m.path == '/tmp/bcompiler/output/master.xlsx'
    assert m.filename == 'master.xlsx'
    assert str(m.quarter) == 'Q2 18/19'
    assert m.year == 2018
    assert len(m.projects) == 3
    assert isinstance(m['PROJECT/PROGRAMME NAME 1'], ProjectData)


def test_master_key_filter(master):
    q1_2017 = Quarter(1, 2017)
    m1 = Master(q1_2017, master)
    for t in m1['PROJECT/PROGRAMME NAME 1'].key_filter('SRO'):
        try:
            assert 'SRO FULL NAME 1' in t[1]
        except (AssertionError, TypeError):
            continue


def test_master_key_filter_missing_key(master):
    q1_2017 = Quarter(1, 2017)
    m1 = Master(q1_2017, master)
    with pytest.raises(KeyError):
        for t in m1['PROJECT/PROGRAMME NAME 1'].key_filter('NOT HERE'):
            continue


def test_pull_iterable_from_master_based_on_key(master):
    p1 = Master(Quarter(1, 2017), master)['PROJECT/PROGRAMME NAME 1']
    assert p1.pull_keys(['SRO Full Name']) == [('SRO Full Name', 'SRO FULL NAME 1')]
    assert p1.pull_keys(['Quarter Joined']) == [('Quarter Joined', 'QUARTER JOINED 1')]
    assert p1.pull_keys(['SRO Full Name', 'Quarter Joined']) == [
        ('SRO Full Name', 'SRO FULL NAME 1'),
        ('Quarter Joined', 'QUARTER JOINED 1'),
    ]
    assert p1.pull_keys(['SRO Full Name', 'Non Existent Key']) == [
        ('SRO Full Name', 'SRO FULL NAME 1'),
    ]


def test_pull_iterable_from_master_based_on_key_flat(master):
    """
    Same as above, but only the value is returned, not the key with it.
    """
    p1 = Master(Quarter(1, 2017), master)['PROJECT/PROGRAMME NAME 1']
    assert p1.pull_keys(['SRO Full Name', 'Quarter Joined'], flat=True) == [
        'SRO FULL NAME 1',
        'QUARTER JOINED 1'
    ]
    assert p1.pull_keys([
        'Working Contact Name',
        'Working Contact Telephone',
        'Working Contact Email'
    ], flat=True) == [
        'WORKING CONTACT NAME 1',
        'WORKING CONTACT TELEPHONE 1',
        'WORKING CONTACT EMAIL 1'
    ]
    assert p1.pull_keys([
        'Working Contact Name',
        'Working Contact Telephone',
        'Non-Existent Key'
    ], flat=True) == [
        'WORKING CONTACT NAME 1',
        'WORKING CONTACT TELEPHONE 1',
    ]


def test_duplicate_keys_in_master(master, caplog):
    m = Master(Quarter(1, 2017), master)
    m.duplicate_keys(True)
    assert "WARNING" in caplog.text
