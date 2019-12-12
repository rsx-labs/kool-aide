# How to Retrieve Data

## Retrieving list of projects

```
c:\your\directory>kool-aide get -m project 
```
The command above will display all projects in tabular format on the screen. To customized the output, several options are provided.

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

Example 4 : Getting all projects and saving it to an excel file with the current date information(i.e. data_January2019.xlsx)
```
c:\your\directory>kool-aide get -m project --output data_[LM][Y].xlsx --format excel
```

Example 5 : Getting first 5 projects, retrieving only the PROJ_ID and PROJ_NAME and displaying it on screen
```
c:\your\directory>kool-aide get -m project --limit 5 --params {\"columns\":[\"PROJ_ID\",\"PROJ_NAME\"]}
```
## Retrieving Commendations in Excel File

A list of Commendations can be retrieved by using the get model and get view option. The get model retrieves raw data from the Commendations table.  The get view option gets its data from the vw_Commendation view and is more suited for reporting purposes.

Retrieving the Excel list of commendations using the get view option

This retrieves all commendations
```
c:\your\directory\>kool-aide get -vw commendation --output c:\file.xlsx --format excel
```

This uses the --params option to retrieve only commendation from 2 projects from a particular month and year
```
c:\your\directory\>kool-aide get -vw commendation --output c:\file.xlsx --format excel --params {\"projects\":[\"Fresco\",\"Savers\"],\"months\":[10,11,12],\"year\":2018}
```
Where:
months : specify the month/s. If empty, the current month is used
year : specify the year. If empty, the current year is used.
projecs : specify the list of projects to retrieve. If empty, all projects.

Note:
All basic params such as 'sorts' and 'columns' is supported. 


## Note
```
The syntax are the same for all supported models and views. However, it is still in development stage, only the following are supported

Models:
- project
- week-range
- employee
- division
- department
- commendation

Views
- status-report
- asset-inventory

```