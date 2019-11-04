# How to Insert Data

Kool-aide currently supports the adding of data on limited tables. Limited business logic are supported but can easily be expanded if necessary. For now, only json file containing an array of data to be inserted is supported. Templates of the different models can be found on the model_templates folder of the repo. 

This function of kool-aide enables other script to batch upload data to AIDE or enables a user to add new data to AIDE via the command prompt.

## Inserting New Project/s to AIDE

1. Create an input file (currently supports json onle file). If you have doubts of the format, use the model template provided. The input file accepts array, so you can insert multiple records at the same time.

```
[
        {
                "PROJ_NAME" : "Test Proj A",
                "CATEGORY" : 1,
                "BILLABILITY" : 0,
                "EMP_ID" : 101010,
                "DSPLY_FLG" : 1,
                "PROJ_CD" : "5554A"
        },
        {
                "PROJ_NAME" : "Test Proj B",
                "CATEGORY" : 1,
                "BILLABILITY" : 0,
                "EMP_ID" : 101010,
                "DSPLY_FLG" : 1,
                "PROJ_CD" : "5554B"
        }
]
```

2. Run kool-aide

```
c:\your\directory>kool-aide add -m project --input c:\Temp\new_proj.json
```

## Note
```
The syntax are the same for all supported models. However, it is still in development stage, only the following are supported

Models:
- project
- week-range
- employee
- division
- department
- commendation
```