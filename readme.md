# Project kool-aide
A nerdy take on a desktop oriented application we call AIDE. Its a (ko)mmand (o)riented (l)ite interface to AIDE. 

## Installation
- Download the executable file to your machine. You may add the exe location to your system path for easier access
- Download the initial kool-aide-setting.json file and modify the user information.
- Fire up your command line and type away!

## Usage
All features can be accessed via command line

### Displaying the help page

```
c:\your\directory>kool-aide -h
```

### Displaying the projects

```
c:\your\directory>kool-aide get -m projects [-l no_of_records][-f result_format]
```

Example : Getting first 5 projects and saving it as an excel file

```
c:\your\directory>kool-aide get -m projects -l 5 -f excel
```

```
c:\your\directory>kool-aide get -m status-report-view -l 100 -p {\"weeks\":[3],\"columns\":[\"Project\",\"Description\"],\"sorts\":[\"Project\"]}

```
## Notes



