
APP_TITLE = 'kool-aide'
APP_VERSION = '0.0.1'

CMD_ACTIONS =[ 
    'create','retrieve','update','delete','generate-report'
]
REPORT_TYPES = [
    'team-monthly-status',
    'team-weekly-status',
    'employee-monthly-status',
    'employee-weekly-status',
    'employee-outstanding-tasks'
]

SUPPORTED_MODELS = [
    'employee', 'attendance'
]



"""
sample :
>kool-aide -a create -m employee --input emp.csv --csv [--output out.json] [--quiet] [-u user@user.com] [-p xxxxxxx]
>kool-aide -a update -m attendance --input attendance.json  [--output out.json] [--quiet] [-u user@user.com] [-p xxxxxxx]
>kool-aide -a create -m employee --input emp.json [--output out.json] [--quiet] [-u user@user.com] [-p xxxxxxx]
>kool-aide -a retrieve -m employee --input "{id:1}" --inline [--output out.json] [--quiet] [-u user@user.com] [-p xxxxxxx]
>kool-aide -a delete -m employee --input "{id:1}" --inline [--output out.json] [--quiet] [-u user@user.com] [-p xxxxxxx]

>kool-aide -a generate-report -r tl-monthly-report --input "{project_id:1;month_id:4}" --inline [--output out.json] [--quiet] [-u user@user.com] [-p xxxxxxx]

"""