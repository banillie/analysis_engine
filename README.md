# analysis_engine 

## About analysis_engine
Collection of code to run portfolio management reporting processes in the Department for Transport. 

Code is still in development, although robustness is improving. `analysis_engine` is the latest complete repository 
for managing the reporting process and still a work in progress. 

Specific framework is required for this code to work. Key requirements are: 
1. Management of data commission and collection via [Datamaps](https://www.datamaps.twentyfoursoftware.com). See more information at website. 
2. Have the correct directories and input files (including master data) set up on your local machine. See below. 

## Directories and file paths
The following directories need to be set up on your local machine. `analysis_engine` is able to handle different operating systems. 

Create the following directories in your `My Documents` directory:

    analysis_engine
    analysis_engine/core_data
    analysis_engine/input
    analysis_engine/output

Master data is stored in the `core_data` directory in the form of individual excel wbs containing master data for each quarter.
Master wbs must be saved in the `core_data` directory with correct title formats 
e.g `(root_path/'core_data/master_3_2019.xlsx', 3, 2019)`
 
`analysis_engine` work by converting these master wbs into python dictionaries, which are then used as the basis for 
running all outputs for `analysis_engine`. See further details on having correct environment. 

Some `analysis_engine` programmes require for there to be a excel document to be present in the `input` folder. In these cases the wb will
be used as a blank/blueprint for the placing of project values/calculations into the required output format. Guidance 
for running each programme will set-up whether an input documents is required. Input files are loaded into 
`analysis_engine` via specifying file paths e.g. `root_path/'input/[name of input file].xlsx`

All outputs from `analysis_engine` will be placed into the `output` folder. Each output file can been named as desired by
the user. Output file names are generated via file paths e.g. `root_path/'output/[name of output file].xlsx`

## Environment

If you are familiar with setting up virtual environments and working with Pycharm this section can be skipped. 

Firstly you need to make sure Python is installed. To install the latest version of Python go to its website and use its
installation wizard. 

Secondly install the community edition of PyCharm from the PyCharm website. 

Thirdly open pycharm. Select create a new project. In the new project window make no changes to the default option and 
select create. 

Fourthly once the default project (probably titled untitled) has been created and select `VCS` `Get from Version 
Control`. Ensure that `Git` is selected in the version control drop down and provide the following URL 
[analysis_engine repo](https://github.com/banillie/analysis_engine). Then click Clone. This will probably prompt a 
Cannot Run Git message. If this happens clink on Download and download the latest release from the git website. 
Install with all the standard/default options. At this point you will likely need to restart your computer. After 
restarting try the `VCS` `Get from Version Control` steps again. This should complete the `analysis_engine` repo 
should appear in PyCharm. 

Fifthly choose `File` `Settings` `Project: analysis_engine` `Project Interpreter`. In the project interpreter drop down
bar select `Show All`. There should be one Project Interpreter displayed. Select and delete this (with - symbol). 
Hit the + symbol and select ok. This will create another Virtual Environment for your local `analysis_engine`. click ok
again until you exist settings. 

Finally open the `Terminal` window in the bottom left tab of the Pycharm window. Ensure that the cursor is flashing 
against the virtual environment for this `analysis_engine`. The file path should start with (Venv). If it doesn't close
the terminal and open it again until the (venv) shows. In the terminal enter the following `pip install -r 
requirements.txt`. 

This should complete your local environment set up. 