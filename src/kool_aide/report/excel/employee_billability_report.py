import json
import pandas as pd
from datetime import datetime
import xlsxwriter
from typing import List

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import get_cell_address, \
    get_cell_range_address, get_version

from kool_aide.model.cli_argument import CliArgument
from kool_aide.assets.resources.messages import *

class EmployeeBillabilityReport:

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
        self._footer_format = None
        self._wrap_content = None
        self._number_two_places = None

        self._log('creating component ...')

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report.employee_billability_report]", level)

    def generate(self, format: str) -> None:
        if format == OUTPUT_FORMAT[3]: # excel
            return self._generate_excel_report()
        else:
            return False, NOT_SUPPORTED

    def _generate_excel_report(self) -> None:
        try:
            self._workbook = self._writer.book
            self._main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            self._sub_header_format = self._workbook.add_format(SHEET_SUB_HEADER)
            self._sub_header_format2 = self._workbook.add_format(SHEET_SUB_HEADER2)
            self._footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            self._wrap_content = self._workbook.add_format(SHEET_CELL_WRAP)
            self._header_format_orange = self._workbook.add_format(SHEET_HEADER_ORANGE)
            self._header_format_gray = self._workbook.add_format(SHEET_HEADER_GRAY)
            self._number_two_places = self._workbook.add_format({'num_format':'0.00'})
            self._report_title = self._workbook.add_format(SHEET_TITLE)
        
            self._create_summary()
            self._create_week_summary()
               
            self._writer.save()
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_summary(self):
        try:
            data_frame = self._data
            sheet_name = 'Summary'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(25,1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(
                current_row, 
                current_col, 
                "Billability Report Summary", 
                self._report_title
            )

            worksheet.set_column(0,0,25,self._cell_wrap_noborder)
            worksheet.set_column(1,50,5,self._cell_wrap_noborder)
            worksheet.set_row(32, 152)     

            current_row += 2

            current_row = 31
     
            worksheet.write(
                current_row, 
                current_col, 
                "Billability Per Employee", 
                self._main_header_format
            )

            worksheet.write(current_row, current_col+1, "", self._main_header_format) 
            worksheet.write(current_row, current_col+2, "", self._main_header_format) 

            current_row += 1

            group_by_employee = data_frame.groupby(['Employee Name'])             
            employee_df = pd.DataFrame(group_by_employee.size().reset_index())
            format_vertical_text = self._sub_header_format
            format_vertical_text.set_rotation(90)
            format_vertical_text.set_align('bottom')
            employee_dict = {}
            start_row = current_row

            worksheet.write(current_row, current_col, "", self._sub_header_format2)
            for index,row in employee_df.iterrows():
                worksheet.write(
                    current_row, 
                    current_col+index+1, 
                    str(row['Employee Name']).strip(), 
                    format_vertical_text
                )

                employee_dict[row['Employee Name'].strip()]=1+index

            current_row +=1

            group_by_project = data_frame.groupby(['Project'])             
            project_df = pd.DataFrame(group_by_project.size().reset_index())
           
            project_dict = {}
            for index,row in project_df.iterrows():
                worksheet.write(current_row, 0, row['Project']) 
                project_dict[row['Project'].strip()]=current_row
                current_row += 1
            
            for col in range(0,len(employee_dict)+1):
                worksheet.write(current_row, current_col+col, "", self._header_format_gray) 
            current_row += 2
            last_row = current_row

            group_per_project_emp = data_frame.groupby(['Project','Employee Name']) 
            for key, values in group_per_project_emp['Hours'].sum().items():
                  worksheet.write(
                      project_dict[key[0].strip()], 
                      employee_dict[key[1].strip()], 
                      values
                    )  
                  current_row += 1
        
            chart = self._writer.book.add_chart({'type':'column', 'subtype':'stacked'}) 
            
            categories_address = get_cell_range_address(
                get_cell_address(1,start_row + 1),
                get_cell_address(len(employee_dict),start_row +1),
                sheet=sheet_name
            )
        
            for key,value in project_dict.items():
                values_address = get_cell_range_address(
                    get_cell_address(1, value +1),
                    get_cell_address(len(employee_dict),value+1),
                    sheet=sheet_name
                )
                
                chart.add_series({
                    'name': f'={sheet_name}!{get_cell_address(0,value+1)}',
                    'categories': f'={categories_address}',
                    'values':     f'={values_address}',
                    'data_labels' : {'value':True}
                })

           
            chart.set_x_axis({'name': 'Employee'}) 
            chart.set_y_axis({'name': 'Hours'}) 
            chart.set_style(10)
            chart.set_title({'name':'Employee Billability Summary'})
            worksheet.insert_chart(
                'A2',
                chart,
                {'x_scale': 2.4, 'y_scale': 2,'x_offset':10,'y_offset':10}
            )

            end_row = last_row
            end_row += 2
            worksheet.write(
                end_row, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                self._footer_format
            )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_week_summary(self):
        try:
            data_frame = self._data
            grouped_per_week =  data_frame.groupby(['Week Range'])
            
            for key, values in grouped_per_week:
                df_per_week = pd.DataFrame(grouped_per_week.get_group(key))

                range_date = datetime.strptime(key,"%Y%m%d")
                range_date_string ='{:02d}'.format(range_date.month) \
                    + '{:02d}'.format(range_date.day)
                
                key_string = f'WeekEnding_{range_date_string}'

                sheet_name = key_string
            
                current_row = 0
                current_col = 0

                worksheet = self._writer.book.add_worksheet(sheet_name)

                title_range = get_cell_range_address(
                    get_cell_address(0,1),
                    get_cell_address(25,1)
                )
                worksheet.merge_range(title_range,'','')

                worksheet.write(
                    current_row, 
                    current_col, 
                    f"Billability Report Summary for {sheet_name}",
                    self._report_title
                )

                worksheet.set_column(0,0,25,self._cell_wrap_noborder)
                worksheet.set_column(1,50,5,self._cell_wrap_noborder)
                worksheet.set_row(32, 152)     

                current_row = 31
        
                worksheet.write(
                    current_row, 
                    current_col, 
                    "Billability Per Employee", 
                    self._main_header_format
                )

                worksheet.write(current_row, current_col+1, "", self._main_header_format) 
                worksheet.write(current_row, current_col+2, "", self._main_header_format) 

                current_row += 1

                group_by_employee = data_frame.groupby(['Employee Name'])             
                employee_df = pd.DataFrame(group_by_employee.size().reset_index())
                format_vertical_text = self._sub_header_format
                format_vertical_text.set_rotation(90)
                format_vertical_text.set_align('bottom')
                employee_dict = {}
                start_row = current_row

                worksheet.write(current_row, current_col, "", self._sub_header_format2)
                for index,row in employee_df.iterrows():
                    worksheet.write(
                        current_row, 
                        current_col+index+1, str(row['Employee Name']).strip(), format_vertical_text
                    )

                    employee_dict[row['Employee Name'].strip()]=1+index

                current_row +=1

                group_by_project = data_frame.groupby(['Project'])             
                project_df = pd.DataFrame(group_by_project.size().reset_index())
            
                project_dict = {}
                for index,row in project_df.iterrows():
                    worksheet.write(current_row, 0, row['Project']) 
                    project_dict[row['Project'].strip()]=current_row
                    current_row += 1
                
                for col in range(0,len(employee_dict)+1):
                    worksheet.write(current_row, current_col+col, "", self._header_format_gray) 
                current_row += 2
                last_row = current_row

                group_per_project_emp = df_per_week.groupby(['Project','Employee Name']) 
                for key, values in group_per_project_emp['Hours'].sum().items():
                    worksheet.write(
                        project_dict[key[0].strip()], 
                        employee_dict[key[1].strip()], 
                        values
                    )  
                    current_row += 1
        
                chart = self._writer.book.add_chart({'type':'column', 'subtype':'stacked'}) 
                
                categories_address = get_cell_range_address(
                    get_cell_address(1,start_row + 1),
                    get_cell_address(len(employee_dict),start_row +1),
                    sheet=sheet_name
                )
                for key,value in project_dict.items():
                    values_address = get_cell_range_address(
                        get_cell_address(1, value +1),
                        get_cell_address(len(employee_dict),value+1),
                        sheet=sheet_name
                    )
                    
                    chart.add_series({
                        'name': f'={sheet_name}!{get_cell_address(0,value+1)}',
                        'categories': f'={categories_address}',
                        'values':     f'={values_address}',
                        'data_labels' : {'value':True}
                    })

                chart.set_x_axis({'name': 'Employee'}) 
                chart.set_y_axis({'name': 'Hours'}) 
                chart.set_style(10)
                chart.set_title({'name':'Employee Billability Summary'})
                worksheet.insert_chart(
                    'A2',
                    chart,
                    {'x_scale': 2.4, 'y_scale': 2,'x_offset':10,'y_offset':10}
                )

            end_row = last_row
            end_row += 2
            worksheet.write(
                    end_row, 
                    0, 
                    f'Report generated : {datetime.now()} by {get_version()}', 
                    self._footer_format
                )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

