# analysis_engine 

## About analysis_engine
Collection of code to run portfolio management reporting processes in the Department for Transport. 

Code is still in development, although robustness is improving. `analysis_engine` is the latest complete repository 
for managing the reporting process and still a work in progress. 

Specific framework is required for this code to work. Key requirements are: 
1. Management of data commission and collection via [Datamaps](https://www.datamaps.twentyfoursoftware.com). See more information at website. 
2. Have the correct directories and input files (including master data) set up on your local machine. See below. 

The following directories need to be set up on your local machine. `analysis_engine` is able to handle different operating systems. 

Create the following directories in your `My Documents` directory:

    analysis_engine
    analysis_engine/core_data
    analysis_engine/input
    analysis_engine/output

Master data set is stored with the `core_data` directory in the form of individual excel wbs containing master data for each quarter.
Master wbs should be saved in the `core_data` directory with the correct title formats. `analysis_engine` workks by
converting these master wbs into python dictionaries, which are then used as the basis for running all outputs for 
`analysis_engine`. See further details on having correct environment. 

Some `analysis_engine` programmes require for there to be a excel document to be present in the `input` folder. In these cases the wb will
be used as a blank/blueprint for the placing of project values/calculations into the required output format. Guidance 
for running each programme will set-up whether an input documents is required. Input files are loaded into 
`analysis_engine` via specifying file paths e.g. `root_path/'input/[name of input file].xlsx`

All outputs from `analysis_engine` will be placed into the `output` folder. Each output file can been named as desired by
the user. Output file names are generated via file paths e.g. `root_path/'output/[name of output file].xlsx`

## Setting up environment