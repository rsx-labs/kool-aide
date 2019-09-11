
# kool-aide/library/constants.py

CMD_ACTIONS =[ 
    'create','get','update','delete','generate-report','time-in'
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

MONTHS = [
    'JAN', 'FEB', 'MAR', 'APR',
    'MAY', 'JUN', 'JUL', 'AUG',
    'SEP', 'OCT', 'NOV', 'DEC'
]

DEFAULT_FILENAME = 'result'
PARAM_COLUMNS = 'columns'
PARAM_SORT = 'sorts'
PARAM_WEEK = 'weeks'
PARAM_PROJECT = 'projects'
# PARAM_FLAGS = 'flags'
