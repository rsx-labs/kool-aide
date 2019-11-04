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