
# kool-aide/library/constants.py

CMD_ACTIONS =[ 
    'add','get','edit','delete','generate-report'
]
REPORT_TYPES = [
    'status-report',
    'asset-inventory'
]

SUPPORTED_MODELS = [
    'employee', 'attendance', 'week-range', 'project', 
    'department', 'division', 'commendation'
]

SUPPORTED_VIEWS = [
    'status-report','asset-inventory','commendation', 'contact-list'
]

OUTPUT_FORMAT = [
    'screen','json', 'csv','excel'
]

LONG_MONTHS = [
    'January', 'February', 'March', 'April',
    'May', 'June', 'July', 'August',
    'September', 'October', 'November', 'December'
]

SHORT_MONTHS = [
    'Jan', 'Feb', 'Mar', 'Apr',
    'May', 'Jun', 'Jul', 'Aug',
    'Sep', 'Oct', 'Nov', 'Dec'
]

DEFAULT_FILENAME = 'result'

PARAM_COLUMNS = 'columns'
PARAM_SORT = 'sorts'
PARAM_WEEK = 'weeks'
PARAM_PROJECT = 'projects'
PARAM_IDS = 'ids'
PARAM_DEPARTMENTS = 'departments'
PARAM_DIVISIONS = 'divisions'
PARAM_START_DATE = 'startdate'
PARAM_END_DATE = 'enddate'
PARAM_MONTHS = 'months'
PARAM_YEAR = 'year'
PARAM_FLAG = 'flag'

# excel formats
SHEET_TOP_HEADER = {
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#CC0000',
    'font_color': '#ffffff',
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
SHEET_HEADER_GAINSBORO = {
    'bold': True,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#DCDCDC',
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
SHEET_CELL_FOOTER = {
    'italic':True,
    'font_color': '#808080',
    'font_size':8
}

# mapping
MAP_VIEW_TO_REPORT ={
    'status-report': 'status-report',
    'asset-inventory': 'asset-inventory'
}
