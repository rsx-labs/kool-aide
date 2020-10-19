import json
import pandas as pd
import numpy as nm
from datetime import datetime
from calendar import monthrange
import xlsxwriter
from typing import List

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import get_cell_address, \
    get_cell_range_address, get_version, get_param_value

from kool_aide.model.cli_argument import CliArgument
from kool_aide.assets.resources.messages import *

class ResourcePlannerReport:

    def __init__(self, logger: CustomLogger, settings: AppSetting, 
                    data: pd.DataFrame, writer = None, 
                    arguments: CliArgument = None) -> None:
        
        self._data = data
        self._settings = settings,
        self._logger = logger

        self._writer = writer
        self._workbook = None
        self._main_header_format = None
        self._header_format_orange = None
        self._header_format_gray = None
        self._cell_wrap_noborder = None
        self._cell_total = None
        self._cell_sub_total = None
        self._footer_format = None
        self._wrap_content = None
        self._number_two_places = None
        self._month = 0
        
        if arguments.parameters is not None:
            try:
                json_parameters = json.loads(arguments.parameters)
                months = get_param_value(PARAM_MONTHS,json_parameters)
                year = get_param_value(PARAM_YEAR, json_parameters)
            except:
                months = None
        elif arguments.auto_mode:
            months =[datetime.now().month]
            year = datetime.now().year
            self._month = datetime.now().month


        if months != None:
            self._month = months[0]
        else:
            self._month = datetime.now().month
        
        if year != None:
            self._year = int(year)
        else:
            self._year = datetime.now().year

        self._log('creating component ...')

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report.resource_planner_report]", level)

    def generate(self, format: str) -> None:
        if self._month == 0:
            return False, MISSING_PARAMETER

        if format == OUTPUT_FORMAT[3]: # excel
            return self._generate_excel_report()
        else:
            return False, NOT_SUPPORTED

    def _generate_excel_report(self) -> None:
        try:
            self._workbook = self._writer.book
            
            self._main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            self._footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            self._wrap_content = self._workbook.add_format(SHEET_CELL_WRAP)
            self._report_title = self._workbook.add_format(SHEET_TITLE)
            self._header_format_lt_gray = self._workbook.add_format(SHEET_HEADER_LT_GRAY)

            self._present = self._workbook.add_format(SHEET_HEADER_LT_GREEN)
            self._vl = self._workbook.add_format(SHEET_HEADER_LT_BLUE)
            self._sl = self._workbook.add_format(SHEET_HEADER_LT_RED)
            self._onsite = self._workbook.add_format(SHEET_HEADER_YELLOW)
            self._holiday = self._workbook.add_format(SHEET_HEADER_LT_PINK)
            self._late = self._workbook.add_format(SHEET_HEADER_LT_GREEN)

            self._sub_header = self._workbook.add_format(SHEET_SUB_HEADER)
            self._set_styles()
   
            self._create_planner()
               
            self._writer.save()
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _set_styles(self):
        self._sub_header.set_align('center')
        self._vl.set_align('left')
        self._sl.set_align('left')
        self._holiday.set_align('left')
        self._late.set_align('left')
        self._onsite.set_align('left')
        self._late.set_font_color('#6D6D6D')
        self._main_header_format.set_align('center')
        self._sub_header.set_align('center')
        self._present.set_align('left')

    def _create_planner(self):
        try:
            color_of_day = {
                1: self._onsite,
                2: self._present,
                3: self._sl,
                4: self._vl,
                5: self._sl,
                6: self._vl,
                7: self._holiday,
                8: self._vl,
                9: self._vl,
                10: self._vl,
                13: self._onsite,
                12: self._vl,
                14: self._onsite,
                11: self._late
            }
            data_frame = self._data
            sheet_name = 'Planner'
            month_name = LONG_MONTHS[self._month -1]
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            worksheet.set_column(0,0,35,self._cell_wrap_noborder)
            worksheet.set_column(1,50,6,self._cell_wrap_noborder)
        
            month_info = monthrange(self._year, self._month)
            month_length = month_info[1]
            start_day = month_info[0]
            
            for day in range(1,month_length + 1):
                worksheet.write(2, day, day, self._sub_header) 
                day_of_week = datetime(self._year, self._month,day).weekday()
                worksheet.write(1, day, DAYS_OF_WEEK[day_of_week], self._main_header_format) 
            
            worksheet.write(2, 0, '', self._sub_header) 
            worksheet.write(1, 0, 'Employee Name', self._main_header_format) 

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(month_info[1],1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(0, 0, f'Resource Planner for {month_name} {self._year}', self._report_title)

            current_row +=3
            group_per_emp = data_frame.groupby(['Employee Name']) 
            employee_df = pd.DataFrame(group_per_emp.size().reset_index())
            employee_dict = {}
            for index, row in employee_df.iterrows():
                worksheet.write(
                    current_row,
                    0, 
                    row['Employee Name']
                ) 
                employee_dict[row['Employee Name']] = current_row
                current_row += 1 
            
            last_row= current_row
            contents = {}

            for index, row in data_frame.iterrows():
                temp_date = datetime.strptime(str(row['Date Entry']),'%Y-%m-%d')
                format_cell = self._present
                #self._log(f"****** temp_date : {temp_date} {row['Employee Name']}")
                temp_address = f"{employee_dict[row['Employee Name']]}:{temp_date.day}"
                self._log(f'temp address = {temp_address}',4)

                try:
                    format_cell = color_of_day[int(row['StatusID'])]
                except Exception as e:
                    self._log(f"error = {e}", 2)
                
                content = row['Short Status']
                if temp_address in contents:
                    content = f"{contents[temp_address]}\\{row['Short Status']}"
                else:
                    contents[temp_address] = row['Short Status']

                worksheet.write(
                    employee_dict[row['Employee Name']],
                    temp_date.day,
                    content,
                    format_cell

                )
            
            end_row = last_row + 2
            worksheet.write(
                end_row, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                self._footer_format
            )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    