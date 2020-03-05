# analysis_engine 

## About analysis_engine :zap: :rocket: :factory:
Collection of code to run portfolio management reporting processes in the Department for Transport. 

`analysis_engine` is still in development, although robustness is improving.

A specific framework is required for code to work. Key requirements are: 
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
e.g `(root_path/'core_data/master_3_2019.xlsx', 3, 2019)`. 
 
`analysis_engine` works by converting these master wbs into variables and lists containing python dictionaries, which 
operate together as a local synthetic database, and then used as the basis for generating analytical outputs. 
See further details on setting up the correct virtual environment. 

Some `analysis_engine` programmes require there to be a excel wb present in the `input` folder. In these cases the wb is 
used as a blueprint for the placing of data values/calculations into a required output format, for example dashboards. 
Input files are loaded into `analysis_engine` via specifying file paths e.g. `root_path/'input/[name_of_input_file].xlsx`

All outputs from `analysis_engine` are placed into the `output` directory. Each output file can been named as desired by
the user. Output file names are generated via file paths e.g. `root_path/'output/[name_of_output_file].xlsx`. 

Each file contains guidance on how it can be run . 

## Virtual environment set up

If you are familiar with setting up virtual environments and working with PyCharm this section can be skipped. The below
steps are set out in a particular way to help a user who is unfamiliar/new to working with Python, PyCharm, Github and 
virtual environments. 

Firstly you need to make sure Python is installed. The most straight forward way to ensure that Python is installed is
to simply install the latest version of Python from its website. Go to the Python website and download the option 
provided on the website. Accept all default download options. 

Secondly install the community edition of PyCharm from the PyCharm website. Make sure you download the community edition.
Accept all default options, but check the create 64 bit desktop app if you can. 

Thirdly open PyCharm. Select create a new project. In the new project window make no changes to the default option and 
select create. This will create a default project call untitled. 

Fourthly select `VCS` `Get from Version Control`. Ensure that `Git` is selected in the version control drop down and provide the following URL 
https://github.com/banillie/analysis_engine. Then click Clone. This will mostly likely prompt a 
`Cannot Run Git` message. If this happens clink on the `Download` option and download the latest release from the git website. 
Install with all the standard/default options. At this point you will likely need to restart your computer. After 
restarting open Pycharm again and try the `VCS` `Get from Version Control` and URL steps again. This should complete 
a cloning of `analysis_engine` code onto your local computer. 

Fifthly choose `File` `Settings` `Project: analysis_engine` `Project Interpreter`. In the project interpreter drop down
bar select `Show All`. There should be one Project Interpreter displayed. Select and delete this with - symbol. 
Hit the + symbol and then select ok (no need to change default options). This will create another Virtual Environment 
for your local `analysis_engine`. click ok again until you exist settings. 

Finally open the `Terminal` window in the bottom left tab of the Pycharm window. Ensure that the cursor is flashing 
against the virtual environment for `analysis_engine`. The file path should start with (venv). If it doesn't close
the terminal and open it again until the (venv) shows. In the terminal enter the following `pip install -r 
requirements.txt`. 

This should complete your local environment set up. 

Note this process creates the necessary framework for `analysis_engine` to run correctly. All `analysis_engine` code is held remotely
on Github. This code is underpinned by having the correct dependencies in your local environment. See the requirements.txt 
file for a list of these dependencies. Key dependencies are Datamaps, Openpyxl and a library of code which handles the 
transformation of master data into a synthetic data base. This code is held here [projectlibrary](https://github.com/banillie/projectlibrary). 
Note that in `projectlibrary` `analysis` `data.py` the master data section shows the file paths used to convert master data into python dictionaries.  

## Maintaining virtual environment

`analysis_engine` is in development and updated regularly to improve performance and add new code. Each time it is used on
 a local machine the user should do the following to ensure that they have the most up-to-date code. 

1. `VCS` `Git` `Pull`. This will pull down all the latest code from the remote repository onto your local machine. If necessary
the user should accept merging with the master.

2. In the terminal type `pip install -U -r requirements.txt` This will update your virtual environment so that you have 
all the latest dependencies. 

