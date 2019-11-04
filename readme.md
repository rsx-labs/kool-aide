# project kool-aide

```                                           
                  ****************************              
                  **                        **              
                  **   &@@.                 **              "a kool and nerdy take on a
                  **    @@@@/               **               desktop oriented application.
                  **      @@@@*             **               its a (ko)mmand (o)riented
                  **     /@@@@              **               (l)ite interface to AIDE"
                  **   %@@@%    **********  **           
                  **   .@@      **********  **           
                  ****************************                                                                                                              
```

## Installation
- Download the executable file to your machine. [TBD]
- Download the initial kool-aide-settings.json file and modify the user information.
- Fire up your command line and type away!

## Usage
All features can be accessed via command line. Initially, only data retrieval, data upload and report generation is supported. Additional data management functions will be enabled once the authentication module is implemented. Let's start drinking!

### Displaying the help page

```
c:\your\directory>kool-aide -h

usage: main.py [-h] [-r {status-report,asset-inventory}]
               [-m {employee,attendance,week-range,project,department,division,position, commendation}]
               [-vw {status-report,asset-inventory}] [--input INPUT_FILE]
               [--output OUTPUT_FILE] [--uid USER_ID] [--password PASSWORD]
               [--quiet] [--interactive] [--format {screen,json,csv,excel}]
               [--limit [RESULT_LIMIT]] [--params PARAMETERS] [--autorun] [-v]
               {add,get,update,delete,generate-report}

A 'fun' interface for AIDE.

positional arguments:
  {create,get,update,delete,generate-report}
                        the action to perform

optional arguments:
  -h, --help            show this help message and exit
  -r {status-report,asset-inventory}, --report {status-report,asset-inventory}
                        report name to generate
  -m {employee,attendance,week-range,project,department,division,position, commendation}, --model {employee,attendance,week-range,project,department,division,position, commendation}
                        the data model to use
  -vw {status-report,asset-inventory}, --view {status-report,asset-inventory}
                        the data view to use
  --input INPUT_FILE    the input file
  --output OUTPUT_FILE  the output file
  --uid USER_ID         the user id
  --password PASSWORD   user password
  --quiet               quiet mode
  --interactive         interactive mode
  --format {screen,json,csv,excel}
                        result format
  --limit [RESULT_LIMIT]
                        limit the number of records
  --params PARAMETERS   command parameters in json
  --autorun             command is run via another application or script
  -v, --version         show program's version number and exit

Oh, yeah!

```

### Displaying the application version

```
c:\your\directory>kool-aide --version

# OUTPUT : kool-aide [DEV] v0.0.2
```
Where:
  - first item refers to the application name, kool-aide
  - second item refers to the current release (i.e. DEV = development release). Other valuse may be BETA, PROD etc...
  - third item is the actual version

### Generating reports
  
Detailed discussion can be found on this link : [How to generate reports](/docs/how_to_generate_report.md)

### Retrieving data

Detailed discussion can be found on this link : [How to retrieve data](/docs/how_to_get_data.md)

### Inserting data

Detailed discussion can be found on this link : [How to insert data](/docs/how_to_add_data.md)

### Updating data

### Deleting Data

### Miscellaneous operations

### Filename generation 

Kool-aide supports configurable file names for its output. Please see below for reference:
- [Y]  : Year
- [M]  : Month in 2 digit notation
- [SM] : Short month notation (JAN, FEB, MAR ...)
- [LM] : Long month notation ( January, February, March ..) 
- [D]  : Day in 2 digit notation

Examples:
```
c:\your\directory>kool-aide get -m project --output projects_asof_[M][D][Y].xlsx --format excel
# OUTPUT FILE : project_asof_09252019.xlsx

c:\your\directory>kool-aide get -m project --output projects_asof_[LM]_[Y].xlsx --format excel
# OUTPUT FILE : project_asof_September_2019.xlsx
```

Note :
```
Some output file formats currently does not support the filename date notation
```

### Configuring kool-aide to run as scheduled script

To be able to call kool-aide from a scheduler, you can:
- direct call to kool-aide with all the needed parameters
- create a batch script that performs the direct call. This way, adding the necessary parameters is less cumbersome

Note :
```
Use the --autorun switch when using kool-aide as part of another script. While the --autorun switch is not yet supported on most calls, it is coming.
```
Example 1 : Generate status report batch file that can be called by the task scheduler
```
@echo off
rem a script that calls kool-aide for report generation
kool-aide generate-report -r status-report --format excel --output "e:\\Temp\\Status Report\\RS_TeamStatusReport_[LM]_[Y].xlsx" --autorun
```

### Log files
When things go wrong, we look at the logs :) Kool-aide logs are created on the same directory where the executable resides. The format is a regular log/txt file, rolling every 1MB of content. The application keeps the last 10 generated files. The amount of information logged is configurable via the kool-aide-setting.json.

## Developer Notes

### [How to Setup Dev Environment](/docs/how_to_setup_devenv.md)

## Release notes

|   Version	|  Changes 	|
|---	|---	|
|  0.0.1 	|  - status report generation, retrieve project list, retrieve week range list	|
|   0.0.2	|  - asset inventory report generation, retrieve employee list, refactoring	|
|   0.0.3	|  - batch upload (employee, project, division, department, commendation) using json input file,  refactoring	|
|   	|   	|