# analysis_engine 

Software for portfolio management reporting and analysis in the UK Department for Transport, operated via command line 
interface (CLI) prompts. 

## Installing
Python must be installed on your computer. If not already installed, it can be installed via the python website
[here](https://www.python.org/downloads/). **IMPORTANT** ensure that `Add Python to PATH` is ticked when provided 
with the option as part of the installation wizard. 

Open the command line terminal (Windows) or bash shell and install via `pip install analysis_engine`.

## Directories, file paths and poppler.
In order to operate the correct directories and files must be set-up and saved on the user's computer. 
`analysis_engine` is able to handle different operating systems. 

Create the following directories in your `My Documents` directory:

    |-- ipdc
        |-- core_data
            |-- json
        |-- input
        |-- output
    |-- top250
        |-- core_data
            |-- json
        |-- input
        |-- output


Each reporting process e.g. ipdc and top250, respective `core_data` directorates require:
1) excel master data files; 
2) excel project information file; and,
3) A confi.ini file. This file lists and master data and project information file names.

As a minimum the `input` folder should have the following documents `summary_temp.docx`, 
`summary_temp_landscape.docx`. In addition `ipdc\input` should have the 
`dashboards_master.xlsx` file. 

All outputs from analysis_engine will be saved into the `output` directory.

The `json` folder is where analysis_engine saves master data in an easily accessible 
format (.json) and after setup can be ignored by the user. 

Unfortunately there is one further manual installation, related to a package within analysis_engine 
which enables high quality rendering of graphical outputs to word documents. On Windows do the following:

1) Download zip of poppler release from this link https://github.com/oschwartz10612/poppler-windows/releases/download/v21.03.0/Release-21.03.0.zip.
2) unzip and move the whole directory to My Documents.
3) Add the poppler bin directory to PATH following these instructions
   https://www.architectryan.com/2018/03/17/add-to-the-path-on-windows-10/
4) Reboot computer.

Mac users should follow instructions here https://pypi.org/project/pdf2image/

Most Linux distributions should not require any manual installation.  

## Operating analysis_engine

To operate analysis_engine the user must enter the initial **_command_** 
`analysis` followed by a _**subcommand**_ to specify the reporting process e.g
`ipdc` or `top250` and then finally an analytical output **_argument_**, the options
for which are set out below. 

**NOTE** the `--help` option is available throughout the entire command
line prompt construction process and the user should use it for guidance on what subcommands
and arguments are available for use. 

analysis_engine currently has the following _arguments_:

`initiate` The user must enter this command
every time master data, contained in the core_data directory, is updated.
The initiate checks and validates the data in a number of ways. 

`dashboards` populates the IPDC PfM report dashboard. A blank template dashboard 
must be saved in the ipdc/input directory. (Not currently available for top250.)

`dandelion` produces the portfolio dandelion info-graphic. 

`costs` produces a cost profile trend graph and data. (Not currently available for top250.)

`milestones` produces milestone schedule graphs and data.

`vfm` produces vfm data. (Not currently available for top250.) 

`summaries` produces project summary reports. 

`risks` produces risk data. (Not currently available for top250.)

`dcas` produces dca data. (Not currently available for top250.)

`speedial` prints out changes in project dca ratings. (Not currently available for top250.)

`query` returns (from master data) specific data required by the user. 

The default for each argument is to return outputs with current and last quarter data.

Further to each argument the user can specify one or many 
further **_optional_arguments_** to alter the analytical output produced. There are 
many optional_arguments available, which vary for each argument, 
and the user should use the `--help` option to specify those that are available. 
