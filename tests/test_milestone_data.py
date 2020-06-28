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
from data_mgmt.data import MilestoneData
from datamaps.api import project_data_from_master
from analysis.data import root_path, bc_index


test_master = project_data_from_master("tests/resources/test_master.xlsx", 4, 2019)
project_names = test_master.projects
master_data = [test_master]


def test_project_names_appear_in_object_project_names_attribute():
    m = MilestoneData(master_data, bc_index, 0)
    project_data = m.project_data(project_names)
    assert "Chutney Bridge.xlsm" in project_data.keys()


def test_baseline_index():
    m = MilestoneData(master_data, bc_index, 0)
    assert isinstance(m.baseline_index, (dict,))
