# analysis_engine 

Software for portfolio management reporting and analysis in the UK Department for Transport, operated via command line 
interface (CLI) prompts. 

## Installing analysis_engine
Python must be installed on your computer. If not already installed, it can be installed via the python website
[here](https://www.python.org/downloads/). **IMPORTANT** ensure that `Add Python to PATH` is ticked when provided 
with the option as part of the installation wizard. 

Open the command line terminal (Windows) or bash shell and install via `pip install analysis_engine`.

## Directories and file paths
The following directories must be set up on your computer. `analysis_engine` is able to handle different operating 
systems. 

Create the following directories in your `My Documents` directory:

    analysis_engine 
    analysis_engine/core_data
    analysis_engine/core_data/data_mgmt
    analysis_engine/core_data/pickle
    analysis_engine/input
    analysis_engine/output

## Operating analysis_engine (ae)

analysis_engine (ae) is operated via the initial command `analysis` followed by the relevant 
subcommands. Subcommands compile the user's desired outputs. All subcommands can be seen
via `analysis --help` and are as follows:

`initiate` The user must enter this command
every time excel master workbook data, saved in the core_data directory, is updated.
Not doing so means ae will continue to use data from the last time initiate was 
used. Ae checks and validates the data in a number of ways, as part of the initiate process 
. See below.

`dashboards` populates the IPDC PfM report dashboard. A blank template dashboard 
must be saved in the analysis/input directory.

`dandelion` produces the portfolio dandelion infographic. Note early version/release.

`costs` produces the cost profile trend graph and data (early version needs more
                        testing).
                   
    milestones          milestone schedule graphs and data (early version
                        needs more testing)
    vfm                 vfm analysis
    summaries           summary reports
    risks               risk analysis
    dcas                dca analysis
    speedial            speed dial analysis
    matrix              cost v schedule chart
    query               return data from core data`


