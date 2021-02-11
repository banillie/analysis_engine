# analysis_engine 

Software for portfolio management reporting and analysis in the UK Department for Transport, operated via command line 
interface (CLI) prompts. 

## Installing
Python must be installed on your computer. If not already installed, it can be installed via the python website
[here](https://www.python.org/downloads/). **IMPORTANT** ensure that `Add Python to PATH` is ticked when provided 
with the option as part of the installation wizard. 

Open the command line terminal (Windows) or bash shell and install via `pip install analysis_engine`.

## Directories and file paths
The following directories must be set up on your computer. `analysis_engine` is able to handle different operating 
systems. 

Create the following directories in your `My Documents` directory:

    |-- analysis_engine
        |--core_data
            |--data_mgmt
            |--pickle
        |--input
        |--output

All excel master data files must be saved in `core_data` using
the correct format e.g. `master_1_2020.xlsx`

The `project_info.xlsx` document must be saved in `core_data/data_mgmt`. 

The following documents should be saved in `input`. `summary_temp.docx` 
`summary_temp_landscape.docx` and `dashboards_master.xlsx`

All outputs from analysis_engine will be saved into the `output` directory.

The `pickle` folder is where analysis_engine saves an easily accessible master data 
set and after setup can be ignored by the user. 

## Operating analysis_engine (ae)

analysis_engine (ae) is operated via the initial **_command_** `analysis` followed by the relevant 
_**subcommand**_. Subcommands compile the user's desired outputs. All subcommands can be seen
via `analysis --help` and are as follows:

`initiate` The user must enter this command
every time excel master workbook data, saved in the core_data directory, is updated.
Not doing so means ae will continue to use data from the last time initiate was 
used. Ae checks and validates the data in a number of ways, as part of the initiate process 
. See below.

`dashboards` populates the IPDC PfM report dashboard. A blank template dashboard 
must be saved in the analysis/input directory.

`dandelion` produces the portfolio dandelion infographic. Note early version/release.

`costs` produces a cost profile trend graph and data.

`milestones` produces milestone schedule graphs and data.

`vfm` produces vfm data. (Note currently no graphs.) 

`summaries` produces project summary reports. 

`risks` produces risk data. (Note currently no graphs.) 

`dcas` produces dca data. (Note currently no graphs.) 

`speedial` prints out changes in project dca ratings. 

`matrix` produces the cost/schedule matrix chart and data. 

`query` returns (from master data) specific data required by the user. 

The default for each subcommand is to return outputs with current and last quarter data.


Further to each subcommand the user has several _**optional arguments**_ available. 
The optional argument available for each subcommand will be shown by `analysis 
[subcommand] --help`. In general the following optional arguments are available
for each subcommand:

`--group` returns output for the project(s) in the specified group. The user can
input either one or a combination of DfT Group name ("HSMRPG", "AMIS", "Rail", "RPE") or 
any number (including one) of the project acronyms e.g. "SARH2".  

`--stage` returns output for the project(s) at the specified planning stage(s).
The user can input either one or a combination of stages ("FBC", "OBC", "SOBC",
"pre-SOBC").

`--quarters` returns output for specified quarter(s). Must be in correct
format e.g. "Q3 20/21". 

`--baselines` returns output for specified baseline(s). Options here are (
"current", "last", "bl_one", "bl_two", "bl_three", "standard", "all"). "current"
and "last" refer to the current and last quarter, so are not true baselines. The
first baseline is therefore "bl_one". The "standard" option will return "current" 
"last" and "bl_one". "all" returns all up to "bl_three".

`--chart` where subcommands automate chart production the user can specify whether to
"show" or "save" the chart. 

`--title` where subcommands automate chart production in some instances the user will
be required to or can chose to provide a title for the chart e.g. "chart title".

###milestone analysis 

For the milestones subcommand there are also the following optional arguments.

`--dates` enables the user to specify dates of interest in the format "start date"
 "end date" e.g. "15/6/2015" "24/3/2020". 

`--dl` means date line and enables the user to include in the graph output a blue 
line to denote a reference date of interest e.g. "1/9/2020". "today" can also be 
entered. 