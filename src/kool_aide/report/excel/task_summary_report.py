import json
import pandas as pd
from datetime import datetime
import xlsxwriter
from typing import List


from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import get_version


from kool_aide.model.cli_argument import CliArgument
from kool_aide.assets.resources.messages import *

class TaskSummaryReport:

    def __init__(self, logger: CustomLogger, settings: AppSetting, 
                    data: pd.DataFrame, writer = None) -> None:
        
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

        self._log('creating component ...')

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report.task_summary_report]", level)

    def generate(self, format: str) -> None:
        if format == OUTPUT_FORMAT[3]: # excel
            return self._generate_excel_report()
        else:
            return False, NOT_SUPPORTED

    def _generate_excel_report(self) -> None:
        try:
            data_frame = self._data
            self._workbook = self._writer.book
            self._main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            self._header_format_orange = self._workbook.add_format(SHEET_HEADER_ORANGE)
            self._header_format_gray = self._workbook.add_format(SHEET_HEADER_GRAY)
            self._cell_wrap_noborder = self._workbook.add_format(SHEET_CELL_WRAP_NOBORDER)
            self._cell_total = self._workbook.add_format(SHEET_HEADER_LT_GREEN)
            self._cell_sub_total = self._workbook.add_format(SHEET_HEADER_GAINSBORO)
            self._footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            # drop_columns = ['WeekRangeStart', 'WeekRangeId', 'ProjectId']
            # column_headers = [
            #     'Project',
            #     'Project Code',
            #     'Rework',
            #     'Ref ID',
            #     'Description',
            #     'Severity',
            #     'Incident Type',
            #     'Assigned Employee',
            #     'Phase',
            #     'Status',
            #     'Start Date',
            #     'Target Date',
            #     'Completion Date',
            #     'Estimate',
            #     'Week Effort',
            #     'Actual Effort',
            #     'Comments',
            #     'Inbound Contacts',
            #     'Week End Date',
            #     'DepartmentID',
            #     'DivisionID'
            # ]

            # data_frame.drop(drop_columns, inplace=True, axis=1)
            # data_frame.columns = column_headers

            self._create_project_report(data_frame)
            self._create_employee_report(data_frame)
            # self._create_summary_report(data_frame)
                          
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            # print(f"the file was saved : {file_name}") 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_project_report(self, data_frame: pd.DataFrame)-> None:
        try:
            
            grouped_per_project = data_frame.groupby('ProjectName')

            sheet_name = 'Task Per Project'

          
            worksheet = self._writer.book.add_worksheet(sheet_name)
            worksheet.set_column(0,0,25)
            worksheet.set_column(1,1,15)
            worksheet.set_column(2,2,45)
            worksheet.set_column(3,3,20)
            worksheet.set_column(4,4,20)
            worksheet.set_column(5,5,20)
            current_col = 0
            current_row = 0

            worksheet.write(current_row, current_col, "Tasks Summary Report", self._main_header_format)       
            worksheet.write(current_row, current_col+1, "", self._main_header_format) 
            worksheet.write(current_row, current_col+2, "", self._main_header_format) 
            worksheet.write(current_row, current_col+3, "", self._main_header_format) 
            worksheet.write(current_row, current_col+4, "", self._main_header_format) 
            worksheet.write(current_row, current_col+5, "", self._main_header_format)       
            current_row += 2

            for key, values in grouped_per_project:
                df_per_group = pd.DataFrame(grouped_per_project.get_group(key))
                df_per_group.sort_values(by=['EmployeeName'], inplace= True)

                worksheet.write(current_row, current_col, key, self._header_format_orange) 
                current_row +=1

                worksheet.write(current_row, current_col, 'Assigned Employee', self._header_format_gray) 
                worksheet.write(current_row, current_col+1, 'Reference ID', self._header_format_gray) 
                worksheet.write(current_row, current_col+2, 'Description', self._header_format_gray) 
                worksheet.write(current_row, current_col+3, 'Incident', self._header_format_gray) 
                worksheet.write(current_row, current_col+4, 'Phase', self._header_format_gray) 
                worksheet.write(current_row, current_col+5, 'Status', self._header_format_gray) 
                current_row += 1

                df_employee_tasks = df_per_group[['EmployeeName','RefID','Description','IncidentType','Phase','TaskStatus']]
               
                issue_count = 0
                for index, row in df_employee_tasks.iterrows():
                    worksheet.write(current_row, current_col, row['EmployeeName'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+1, row['RefID'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+2, row['Description'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+3, row['IncidentType'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+4, row['Phase'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+5, row['TaskStatus'], self._cell_wrap_noborder) 
                    issue_count += 1
                    current_row += 1

                worksheet.write(current_row, current_col, f'Issue Count : {issue_count}', self._header_format_gray) 
                worksheet.write(current_row, current_col+1, '', self._header_format_gray) 
                worksheet.write(current_row, current_col+2, '', self._header_format_gray) 
                worksheet.write(current_row, current_col+3, '', self._header_format_gray) 
                worksheet.write(current_row, current_col+4, '', self._header_format_gray) 
                worksheet.write(current_row, current_col+5, '', self._header_format_gray) 
                current_row += 2

                worksheet.write(
                    current_row, 
                    0, 
                    f'Report generated : {datetime.now()} by {get_version()}', 
                    self._footer_format
                )

          
        except Exception as ex:
            self._log(f'error generating main report. {str(ex)}', 2)

    def _create_employee_report(self, data_frame: pd.DataFrame)-> None:
        try:
            
            grouped_per_project = data_frame.groupby('EmployeeName')

            sheet_name = 'Task Per Employee'

          
            worksheet = self._writer.book.add_worksheet(sheet_name)
            worksheet.set_column(0,0,25)
            worksheet.set_column(1,1,15)
            worksheet.set_column(2,2,45)
            worksheet.set_column(3,3,20)
            worksheet.set_column(4,4,20)
            worksheet.set_column(5,5,20)
            current_col = 0
            current_row = 0

            worksheet.write(current_row, current_col, "Tasks Summary Report", self._main_header_format)  
            worksheet.write(current_row, current_col+1, "", self._main_header_format) 
            worksheet.write(current_row, current_col+2, "", self._main_header_format) 
            worksheet.write(current_row, current_col+3, "", self._main_header_format) 
            worksheet.write(current_row, current_col+4, "", self._main_header_format) 
            worksheet.write(current_row, current_col+5, "", self._main_header_format)      
            current_row += 2

            for key, values in grouped_per_project:
                df_per_group = pd.DataFrame(grouped_per_project.get_group(key))
                df_per_group.sort_values(by=['ProjectName'], inplace= True)

                worksheet.write(current_row, current_col, key, self._header_format_orange) 
                current_row +=1

                worksheet.write(current_row, current_col, 'Project', self._header_format_gray) 
                worksheet.write(current_row, current_col+1, 'Reference ID', self._header_format_gray) 
                worksheet.write(current_row, current_col+2, 'Description', self._header_format_gray) 
                worksheet.write(current_row, current_col+3, 'Incident', self._header_format_gray) 
                worksheet.write(current_row, current_col+4, 'Phase', self._header_format_gray) 
                worksheet.write(current_row, current_col+5, 'Status', self._header_format_gray) 
                current_row += 1

                df_project_tasks = df_per_group[['ProjectName','RefID','Description','IncidentType','Phase','TaskStatus']]
                
                issue_count = 0
                for index, row in df_project_tasks.iterrows():
                    worksheet.write(current_row, current_col, row['ProjectName'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+1, row['RefID'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+2, row['Description'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+3, row['IncidentType'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+4, row['Phase'], self._cell_wrap_noborder) 
                    worksheet.write(current_row, current_col+5, row['TaskStatus'], self._cell_wrap_noborder) 
                    issue_count += 1
                    current_row += 1

                worksheet.write(current_row, current_col, f'Issue Count : {issue_count}', self._header_format_gray) 
                worksheet.write(current_row, current_col+1, '', self._header_format_gray) 
                worksheet.write(current_row, current_col+2, '', self._header_format_gray) 
                worksheet.write(current_row, current_col+3, '', self._header_format_gray) 
                worksheet.write(current_row, current_col+4, '', self._header_format_gray) 
                worksheet.write(current_row, current_col+5, '', self._header_format_gray)
                current_row += 2

                worksheet.write(
                    current_row, 
                    0, 
                    f'Report generated : {datetime.now()} by {get_version()}', 
                    self._footer_format
                )

          
        except Exception as ex:
            self._log(f'error generating main report. {str(ex)}', 2)
