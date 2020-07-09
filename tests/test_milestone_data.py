"""
Example of how you can test the MilstoneData class.

## Changes improving modularity

I have moved the MilstoneData class into its own file at data_mgmt.data. This
allows the class to be more easily imported into other modules or test modules
like this one without running lots of extraneous code you don't need.

Remember, designing code with classes serves a few important functions:
    * modularity (splitting up code into clean, separate files and packages),
    * encapsulation (hiding data from the user and providing then with a
    simple interface),
    * testability (the above makes your code more testable, which is good
    thing)

## Test fixtures

When your tests rely on data to run, this is called a 'fixture'. In your case,
the master_4_2019.xlsx file referenced here is your test fixture. The idea with
your fixtures is that they should provide the minimum amount of data needed to
test the code you need.  In this case, you want a full master, because your
testing code that processes data from a master sheet to provide individual
master data.

BUT! You never use real data in your test! You want your fixtures to travel
alongside your test code, i.e. in your Github repo and you should never put
real data in there for confidentiality, etc.

So, you have a problem to solve. You need to replace "master_4_2019.xlsx" in
this here with a dummy master, which is exactly the same format, but contains
nonsense data. I have dozens of such files for bcompiler_engine - check out
https://github.com/hammerheadlemon/bcompiler-engine/tree/master/tests/resources

I've taken the liberty of using one of my fake master files and added it to
this repo at tests/resources/test_master.data.

I've also adapted the test code below to use this master and to run assert
statements on the data found within it. So, the `project_data_from_master()`
function now targets the test_master.xlsx file, which means the data in 
`project_names` has changed.  With this, we can now test our MilstoneData
class to make sure it does what we expect it to.

## Allowing tests to run fast

Also, tests should run fast. Ideally the whole suit of tests for an
application should run in few seconds, although this may differ depending on
the kinds of test being run.

When pytest runs, it "collects" all test functions, by running through the
codebase looking for funcs that are named "test_*". To do this, it seemingly
has to import every file. If you have code at the global level of your modules,
i.e. code that does not sit inside a function or class, this is going to be
run when this process happens. That is why, I think, running pytest can take
longer than it should. To fix this, you should move all your code inside 
functions.

## Running pytest

Install pytest with `pip install pytest` in your virtualenv. I have updated
your requirements.txt file too.

`pytest -v --tb=short --disable-warnings` (-v means one level of verbosity and
--tb=short means that the tracebacks it sends when it hits an error are
shorter and easier to read. You also do not need warnings.).

Go and write some more tests!
"""
from data_mgmt.data import MilestoneData, MilestoneChartData, Masters, CostData
from datamaps.api import project_data_from_master
from analysis.data import list_of_masters_all, a303, hs2_programme, rail_franchising, sarh2
import pytest
import datetime

@pytest.fixture
def abbreviations():
    return {'2nd Generation UK Search and Rescue Aviation': 'SARH2',
                 'A12 Chelmsford to A120 widening': 'A12',
                 'A14 Cambridge to Huntingdon Improvement Scheme': 'A14',
                 'A303 Amesbury to Berwick Down': 'A303',
                 'A358 Taunton to Southfields Dualling': 'A358',
                 'A417 Air Balloon': 'A417',
                 'A428 Black Cat to Caxton Gibbet': 'A428',
                 'A66 Full Scheme': 'A66',
                 'Crossrail Programme': 'Crossrail',
                 'East Coast Digital Programme': 'ECDP',
                 'East Coast Mainline Programme': 'ECMP',
                 'East West Rail Programme (Central Section)': 'EWR (Central)',
                 'East West Rail Programme (Western Section)': 'EWR (Western',
                 'Future Theory Test Service (FTTS)': 'FTTS',
                 'Great Western Route Modernisation (GWRM) including electrification': 'GWRM',
                 'Heathrow Expansion': 'HEP',
                 'Hexagon': 'Hexagon',
                 'High Speed Rail Programme (HS2)': 'HS2 Prog',
                 'HS2 Phase 2b': 'HS2 2b',
                 'HS2 Phase1': 'HS2 1',
                 'HS2 Phase2a':'HS2 2a',
                 'Integrated and Smart Ticketing - creating an account based back office': 'IST',
                 'Intercity Express Programme': 'IEP',
                 'Lower Thames Crossing': 'LTC',
                 'M4 Junctions 3 to 12 Smart Motorway': 'M4',
                 'Manchester North West Quadrant': 'MNWQ',
                 'Midland Main Line Programme': 'MML Prog',
                 'Midlands Rail Hub': 'Mid Rail Hub',
                 'North Western Electrification': 'NWE',
                 'Northern Powerhouse Rail': 'NPR',
                 'Oxford-Cambridge Expressway': 'Ox-Cam Expressway',
                 'Rail Franchising Programme': 'Rail Franchising',
                 'South West Route Capacity': 'SWRC',
                 'Thameslink Programme': 'Thameslink',
                 'Transpennine Route Upgrade (TRU)': 'TRU',
                 'Western Rail Link to Heathrow': 'WRlTH'}

start_date = datetime.date(2020, 6, 1)
end_date = datetime.date(2022, 6, 30)

# test_master = project_data_from_master("tests/resources/test_master.xlsx", 4, 2019)
# project_names = test_master.projects
# master_data = [test_master]

test_masters = list_of_masters_all[1:]
project_names = list_of_masters_all[1].projects
project_names.remove(hs2_programme)
project_names.remove(rail_franchising)
#project_names.remove(sarh2)


@pytest.fixture
def mst():
    return Masters(test_masters, project_names)

def test_Masters_get_baseline_data(mst):
    mst.get_baseline_data('Re-baseline IPDC milestones')
    assert isinstance(mst.bl_index, (dict,))

def test_MilestoneData_project_dict_returns_dict(mst, abbreviations):
    mst.get_baseline_data("Re-baseline IPDC milestones")
    m = MilestoneData(mst, abbreviations)
    assert isinstance(m.project_current, (dict,))

def test_MilestoneData_group_dict_returns_dict(mst, abbreviations):
    mst.get_baseline_data('Re-baseline IPDC milestones')
    m = MilestoneData(mst, abbreviations)
    assert isinstance(m.group_current, (dict,))

def test_MilestoneChartData_group_chart_returns_list(mst, abbreviations):
    mst.get_baseline_data('Re-baseline IPDC milestones')
    m = MilestoneData(mst, abbreviations)
    mcd = MilestoneChartData(milestone_data_object=m)
    assert isinstance(mcd.group_current_tds, (list,))

def test_MilestoneChartData_group_chart_filter_in_works(mst, abbreviations):
    assurance = ['Gateway', 'SGAR', 'Red', 'Review']
    mst.get_baseline_data('Re-baseline IPDC milestones')
    m = MilestoneData(mst, abbreviations)
    mcd = MilestoneChartData(m, keys_of_interest=assurance)
    assert any("Gateway" in s for s in mcd.group_keys)
    assert any("SGAR" in s for s in mcd.group_keys)
    assert any("Red" in s for s in mcd.group_keys)
    assert any("Review" in s for s in mcd.group_keys)



# def test_MilestoneChartData_group_chart_filter_out_works(abbreviations):
#     mst = Masters(list_of_masters_all[1:], list_of_masters_all[1].projects)
#     assurance = ['Gateway', 'SGAR', 'Red', 'Review']
#     m = MilestoneData(mst, abbreviations)
#     mcd = MilestoneChartData(milestone_data_object=m,
#                              keys_of_interest=None,
#                              keys_not_of_interest=assurance,
#                              filter_start_date=start_date,
#                              filter_end_date=end_date)
#     assert 'Gateway' not in mcd.group_keys
#     assert 'SGAR' not in mcd.group_keys

def test_CostData_get_financial_totals_returning_totals(mst):
    mst = Masters(test_masters, project_names)
    mst.get_baseline_data('Re-baseline IPDC cost')
    c = CostData(mst)
    assert isinstance(c.last, (list,))
