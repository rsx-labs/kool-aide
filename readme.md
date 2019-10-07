# project kool-aide
A fun and nerdy take on a desktop oriented application we call AIDE. Its a (ko)mmand (o)riented (l)ite interface to AIDE. 
```

                                                                                                                       
                                                                         /@@@@#                                        
                                                                       /@@@@@@@@&.                                     
                                                                      @@@@@@@@@@@@@.                                   
                                                                    @@@@@@@@@@@@@@@@@%                                 
                                                                  (@@@*   %@@@@@@@@@@@@%                               
                                                                (@@@*       /@@@@@@@@@@@@@.                            
                                                               @@@@           .@@@@@@@@@@@@@/                          
                                                             &@@@.      ..,      #@@@@@@@@@@@@%                        
                                                           #@@@,      /@@@@        #@@@@@@@@@@@@@                      
                                                         ,@@@&        ,@@@@          *@@@@@@@@@@@@@                    
                                                        @@@&          .@@@@            .@@@@@@@. @@@@(                 
                                                      %@@@,            @@@@,              &@@@(    @@@@(               
                                                    %@@@,     .@@@@@@@@@@@@*                *@@@@. *@@@@@(             
                                                   /@@&        @@@@@@@@@@@@(                  *@@@@@@@@@@#             
                               &@@@@@@@             &@@@(      .                                 &@@@@@@.              
                                                      #@@@%               .@#                     (@@@*                
                                                 *%%/   *@@@@,          .@@@@@@.                ,@@@#                  
                                               %@@@@@@@   .@@@@,        &@@@@@@@@/             @@@&                    
                                            /@@@@@@@@@@@@,   #@@@&        .@@@@@@@@%         &@@@.                     
                       @@@@@@@@           #@@@@@@@@@@@@@@@@#   #@@@&        .@@@@@@@@@     #@@@*                       
                                       ,@@@@@@@@@@@@@@@@@@@@@@.  .@@@@*        %@@@@@%   *@@@#                         
                                     .@@@@@@@@@@@*  (@@@@@@@@@@@.  .@@@@*        /@@   .@@@&                           
                                     &@@@@@@@@@     @@@@@@@@@@@@@@%   (@@@&           @@@@                             
                                      @@@@@@(     %@@@@@@@@@@@@@@@@@    (@@@&.      %@@@,                              
                                                @@@@@@@@@@@@@@@@@@@@/      @@@@/  *@@@(                                
                                              &@@@@@@@@@@@@@@@@@@@@@/        %@@@@@@%                                  
                           ,&&&&&&&         @@@@@@@@@@@@@@@@@@@@@@@@/          *@@%                                    
                                          &@@@@@@@@@@@@@@@@@@@@@@@@@/                                                  
                                        &@@@@@@@@@@@@@@@@@@@#@@@@@@@/                                                  
                                     /@  (@@@@@@@@@@@@@@@@&  @@@@@@@@@@@@@@@@@@@@@,                                    
                                   #@@@@%  #@@@@@@@@@@@@(    @@@@@@@@@@@@@@@@@@@@@@(                                   
                                 %@@@@@@@@*  @@@@@@@@@@      @@@@@@@@@@@@@@@@@@@@@@*                                   
                               &@@@@@@@@@@,   .@@@@@@@@@%     @@@@@@@@@@@@@@@@@@@&                                     
                             &@@@@@@@@@@,       &@@@@@@@@@.                                                            
                          .@@@@@@@@@@@.           @@@@@@@@@&                                                           
                         @@@@@@@@@@@               *@@@@@@@@@,                                                         
                      ,@@@@@@@@@@&         ,%@@@@@@@@@@@@@@@@@.                                                        
                      @@@@@@@@@@       %@@@@@@@@@@@@@@@@@@@@@@         .%%%%%%%                                        
                      @@@@@@@#        %@@@@@@@@@@@@@@@@@@@@&.                                                          
                        %@&/          ,@@@@@@@@@@@@&(.                                                                 
                                        %@@@(.                                                                         
                                                                                                                                                                                                                                                                                                                                                       
```

## Installation
- Download the executable file to your machine. [TBD]
- Download the initial kool-aide-settings.json file and modify the user information.
- Fire up your command line and type away!

## Usage
All features can be accessed via command line. Initially, only data retrieval and report generation is supported. Additional data management functions will be enabled once the authentication module is implemented.

### Displaying the help page

```
c:\your\directory>kool-aide -h

usage: main.py [-h] [-r {status-report,asset-inventory}]
               [-m {employee,attendance,week-range,project,department,division,position}]
               [-vw {status-report,asset-inventory}] [--input INPUT_FILE]
               [--output OUTPUT_FILE] [--uid USER_ID] [--password PASSWORD]
               [--quiet] [--interactive] [--format {screen,json,csv,excel}]
               [--limit [RESULT_LIMIT]] [--params PARAMETERS] [--autorun] [-v]
               {create,get,update,delete,generate-report}

A 'fun' interface for AIDE.

positional arguments:
  {create,get,update,delete,generate-report}
                        the action to perform

optional arguments:
  -h, --help            show this help message and exit
  -r {status-report,asset-inventory}, --report {status-report,asset-inventory}
                        report name to generate
  -m {employee,attendance,week-range,project,department,division,position}, --model {employee,attendance,week-range,project,department,division,position}
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

Team Status Report
```
c:\your\directory>kool-aide generate-report -r status-report --format excel --output c:\outpu\dir\status_report_[LM][Y].xlsx --autorun
```
The command above wil generate a status report in excel format on the directory and file name specified. The --autorun switch tells the application that it is being run as an automated script, thus, some settings will be read on the config file. 

Asset Inventory Report
```
c:\your\directory>kool-aide generate-report -r asset-inventory --format excel --output c:\outpu\dir\asset_inventory_[LM][Y].xlsx --autorun
```

### Displaying the projects

```
c:\your\directory>kool-aide get -m project 
```
The command above will display all projects in tabular format on te screen. To customized the output, several options are provided.

Example 1 : Getting first 5 projects. 

```
c:\your\directory>kool-aide get -m project --limit 5
```

Example 2 : Getting the first 10 projects sorted by project name 
```
c:\your\directory>kool-aide get -m project --limit 10 --params {\"sorts\":[\"PROJ_NAME\"]}
```

Example 3 : Getting all projects and saving it to a json file (i.e. data.json)
```
c:\your\directory>kool-aide get -m project --output data.json --format json
```

Example 3 : Getting all projects and saving it to an excel file with the current date information(i.e. data_January2019.xlsx)
```
c:\your\directory>kool-aide get -m project --output data_[LM][Y].xlsx --format excel
```

Note:
```
The syntax are the same for all supported models and views. However, it is still in development stage, only the following are supported

Models:
- project
- week-range
- employee

Views
- status-report
- asset-inventory

```

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
When things go wrong, we look at the logs :) Kool-aide logs are created the the same directory where the executable resides. The format is a regular log file, rolling every 1MB of content. The application keeps the last 10 generated files. The amount of information logged is configurable via the kool-aide-setting.json.


### Release notes
- v0.0.1 TBD
- v0.0.2 TBD
