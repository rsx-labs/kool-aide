# how to setup the dev environment for kool-aide

## pre-req's
- python 3.7+
- ms sql server (express|full)
- the latest backup for the database (with test data)
- your favorite editor (note: initial code is currently developed using vscode)

## basic setup
- clone the repository from github
- create a virtual environment. please see separate section for the instructions.
- activate the virtual environment
- install the required packages listed in requirements.txt via 'pip'
- open the project on your favorite editor. you can use pycharm, visual studio with python support, visual studio code (preferred), atom, sublime, vim, notepad etc.
- restore the database backup

## how to create virtual environment for python 3
make sure you create a virtual environment for the project. numerous tutorials are available in the web, google your hearts out.

## how to install packages from the requirements.txt using 'pip'
to install all dependencies, we need to install all packages fro the requirements.txt.  make sure you have activate the virtual environment before doing the following steps:
- go to the directory where the requirements.txt resides.
- from the terminal, run following command
```
pip install -r requirements.txt
```

note: if you have installed additional new packages, make sure you update the requirements.txt and check-in the changes. to update the requirements.txt, use the following command:
```
pip list > requirements.txt
```

## how to setup the python debugger for visual studio code
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
    "args": ["get", "-m", "project","--limit',"100"]
}

```

## how to build the executable
you can run the codes directly using the command
```
C:\Your\Directory\kool-aide-master\src>python -m kool_aide.cli.main [arguments here...]
```

or, build the executable using pyinstaller. you need to install pyinstaller on your base python installation.
```
pip install pyinstaller
```
run the make-exe.bat from the toolscript folder
```
c:\your\directory\kool-aide-master\toolscripts>make-exe.bat kool-aide [version]
```
note: the version is not currently used

