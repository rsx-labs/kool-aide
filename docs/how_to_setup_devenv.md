# How to setup the dev environment for kool-aide

## Basic setup
- clone the repository from github
- create a virtual environment. please see separate section for the instructions.
- activate the virtual environment
- install the required packages listed in requirements.txt via 'pip'
- open the project on your favorite editor. you can use pycharm, visual studio with python support, visual studio code (preferred), atom, sublime, vim, notepad etc.

## How to create virtual environment for python 3

## How to install packages from the requirements.txt using 'pip'

## How to setup the python debugger for visual studio code
- go to the main folder of the project
- activate the virtual environment
- open the editor in the command line :

```

c:\To\Your\Directory> code . 

```

- create a debugging configuration by editing the projects launch.json file. see sample below. you can create as many configurations you need.

basic debugging setup
``` 

{
    "name": "Debug cli",
    "type": "python",
    "request": "launch",
    "module": "kool-aide.cli.main",
    "cwd":"${workspaceFolder}\\src"
}

```

debugging kool-aide with parameter. i.e. get the status-report-view , limit it to 100 recordes and write it to an excel file

```

{
    "name": "Debug cli",
    "type": "python",
    "request": "launch",
    "module": "kool-aide.cli.main",
    "cwd":"${workspaceFolder}\\src",
    "args": ["get", "-m", "status-report-view","-l',"100","-f","excel"]
}

```


