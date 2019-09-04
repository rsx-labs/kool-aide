
# kool-aide/library/constants.py

CMD_ACTIONS =[ 
    'create','get','update','delete','generate-report'
]
REPORT_TYPES = [
    'team-status-report',
    'status-report',
    'outstanding-tasks'
]

SUPPORTED_MODELS = [
    'employee', 'attendance', 'week-range', 'project', 'status-report-view'
]

DISPLAY_FORMAT = [
    'screen','json', 'csv','excel'
]

DEFAULT_FILENAME = 'result'
PARAM_COLUMNS = 'column'
PARAM_SORT = 'sort'
PARAM_WEEK = 'week'
PARAM_PROJECT = 'project'