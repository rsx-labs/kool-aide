
# kool-aide/library/constants.py

CMD_ACTIONS =[ 
    'add','get','edit','delete','generate-report'
]
REPORT_TYPES = [
    'status-report','asset-inventory',
    'task-report', 'project-billability',
    'employee-billability', 'non-billables',
    'kpi-summary','late-tracking','skills-matrix',
    'resource-planner'
]

SUPPORTED_MODELS = [
    'employee', 'attendance', 'week-range', 'project', 
    'department', 'division', 'commendation'
]

SUPPORTED_VIEWS = [
    'status-report','asset-inventory','commendation', 'contact-list', 
    'leave-summary', 'task', 'action-list', 'lesson-learnt',
    'project-billability', 'employee-billability', 'concern-list',
    'success-register', 'comcell-schedule', 'kpi-summary', 'attendance',
    'skills', 'resource-planner'
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

FISCAL_MONTHS = [
    'April', 'May', 'June', 'July', 'August',
    'September', 'October', 'November', 
    'December','January', 'February', 'March'
]

DAYS_OF_WEEK = [
    'MON','TUE','WED','THU','FRI','SAT','SUN'
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
PARAM_FYS = 'fys'
PARAM_TYPES = 'types'
PARAM_PHASES = 'phases'
PARAM_STATUS = 'status'

# excel formats

SHEET_TITLE = {
    'bold': True,
    'valign': 'top',
    'fg_color': '#000005',
    'font_color': '#ffffff',
    'border': 0,
    'font_size':17
}

SHEET_TOP_HEADER = {
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#6B0000',
    'font_color': '#ffffff',
    'border': 0
}

SHEET_TOP_HEADER_NO_WRAP = {
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#6B0000',
    'font_color': '#ffffff',
    'border': 0
}

SHEET_SUB_HEADER_NO_WRAP = {
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#E9967A',
    'font_color': '#000000',
    'border': 0
}

SHEET_SUB_HEADER = {
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#E9967A',
    'font_color': '#000000',
    'border': 0
}

SHEET_SUB_HEADER_NO_WRAP = {
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#E9967A',
    'font_color': '#000000',
    'border': 0
}
SHEET_SUB_HEADER2 = {
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#DA2809',
    'font_color': '#ffffff',
    'border': 0
}

SHEET_SUB_HEADER2_NO_WRAP = {
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#DA2809',
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

SHEET_HEADER_LT_GRAY = {
    'bold': False,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#EEEDEA',
    'border': 0
}

SHEET_HEADER_LT_GREEN = {
    'bold': False,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#00e600',
    'border': 0,
    'align':'center'
}
SHEET_HEADER_LT_BLUE = {
    'bold': False,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#4dd2ff',
    'border': 0,
    'align':'center'
}
SHEET_HEADER_LT_ORANGE = {
    'bold': False,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#ffaa00',
    'border': 0,
    'align':'center'
}
SHEET_HEADER_LT_RED = {
    'bold': False,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#EC3305',
    'border': 0,
    'align':'center'
}
SHEET_HEADER_LT_PINK = {
    'bold': False,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#ffb3ff',
    'border': 0,
    'align':'center'
}
SHEET_HEADER_YELLOW = {
    'bold': False,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#ffff00',
    'border': 0,
    'align':'center'
}
SHEET_HEADER_GAINSBORO = {
    'bold': True,
    'text_wrap':False,
    'valign': 'top',
    'fg_color': '#DCDCDC',
    'border': 0
}

SHEET_CELL_WRAP = {
    'text_wrap':True,
    'valign':'top'
}

SHEET_CELL_WRAP_NOBORDER = {
    'text_wrap':True,
    'border':0,
    'valign':'top'
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

# mapping  report:view
MAP_VIEW_TO_REPORT ={
    'status-report': 'status-report',
    'asset-inventory': 'asset-inventory',
    'task-report': 'task',
    'project-billability': 'project-billability',
    'employee-billability':'employee-billability',
    'non-billables':'employee-billability',
    'kpi-summary':'kpi-summary',
    'late-tracking':'attendance',
    'skills-matrix':'skills',
    'resource-planner':'resource-planner'
}

# view in excel that are generic
GENERIC_EXCEL_VIEWS = [
    SUPPORTED_VIEWS[2], SUPPORTED_VIEWS[3], SUPPORTED_VIEWS[4],
    SUPPORTED_VIEWS[5], SUPPORTED_VIEWS[6], SUPPORTED_VIEWS[7],
    SUPPORTED_VIEWS[8], SUPPORTED_VIEWS[9], SUPPORTED_VIEWS[10],
    SUPPORTED_VIEWS[11], SUPPORTED_VIEWS[12], SUPPORTED_VIEWS[13],
    SUPPORTED_VIEWS[14], SUPPORTED_VIEWS[15], SUPPORTED_VIEWS[16]
]
