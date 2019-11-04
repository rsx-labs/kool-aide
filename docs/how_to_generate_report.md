# How to Generate Reports

Kool-aide uses Microsoft Excel file (xlsx) as output format (as of now) of the generated report. The list of reports that can be generated is listed on the applications help page ( using the -h switch).

## Team Status Report

```
c:\your\directory>kool-aide generate-report -r status-report --format excel --output c:\outpu\dir\status_report_[LM][Y].xlsx --autorun
```

The command above wil generate a status report in excel format on the directory and file name specified. The --autorun switch tells the application that it is being run as an automated script, thus, some settings will be read on the config file.

## Asset Inventory Report

```
c:\your\directory>kool-aide generate-report -r asset-inventory --format excel --output c:\outpu\dir\asset_inventory_[LM][Y].xlsx --autorun
```

The command above generates the asset inventory report. This command is meant to be run as an automated script as indicated by the --autorun switch.

## Important Parameters

- [ generate-report ] : this is the method you want to invoke. in this case, generate-report.
- [ -r ] : the switch takes the report type to generate. in the example above, the report type is either the stats-report or the asset-inventory.
- [ --format ] : the output format. currently, this is required bu it may become optional and defaulting to 'excel' if none is provided.
- [ --output ] : the file name of the report
- [ --autorun ] : a switch indicating that the command is to be called by another script or program and that some parameters will be read from the config file.

## Note
Reports are different from the raw excel file that can be generated using the 'get' method. The reports contains formatting and computations that the raw file does not have.
