
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

# excel formats
SHEET_TOP_HEADER = {
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 0
}
SHEET_HEADER_ORANGE ={
    'bold': True,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#FDBF42',
    'border': 0
}

SHEET_HEADER_GRAY = {
    'bold': True,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#BBB9B5',
    'border': 0
}
SHEET_HEADER_LT_GREEN = {
    'bold': True,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#66FF00',
    'border': 0
}
SHEET_HEADER_LT_BLUE = {
    'bold': True,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#66FFCC',
    'border': 0
}

SHEET_CELL_WRAP = {
    'text_wrap':True
}

SHEET_CELL_WRAP_NOBORDER = {
    'text_wrap':True,
    'border':0
}
SHEET_CELL_WRAP_NOBORDER_ALT = {
    'text_wrap':True,
    'border':0,
    'fg_color': '#E8E8E8',
}


