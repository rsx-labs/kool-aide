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

class NonBillablesReport:

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

        self._start_employee_data_row = 0
        self._start_weekl_data_row =0
        self._start_week_total_row = 0
        self._start_total_row = 0

        self._end_employee_data_row = 0
        self._end_weekl_data_row =0
        self._end_week_total_row = 0
        self._end_total_row = 0

        self._log('creating component ...')

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report.non_billable_report]", level)

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
            self._report_title = self._workbook.add_format(SHEET_TITLE)
            self._footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            self._wrap_content = self._workbook.add_format(SHEET_CELL_WRAP)
            self._header_format_orange = self._workbook.add_format(SHEET_HEADER_ORANGE)
            self._header_format_gray = self._workbook.add_format(SHEET_HEADER_GRAY)
            self._number_two_places = self._workbook.add_format({'num_format':'0.00'})

            self._sub_header_format.set_align('vcenter')

            self._create_data_summary()
                      
            self._writer.save()
          
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_data_summary(self):
        try:
            data_frame = self._data

            group_by_employee = data_frame.groupby(['Employee Name'])             
            employee_df = pd.DataFrame(group_by_employee.size().reset_index())

            data_frame = data_frame[data_frame['IsBillable'] == 0]
            sheet_name = 'Data'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(15,1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(
                current_row, 
                current_col, 
                "Non Billable Report Summary", 
                self._report_title
            )

            worksheet.set_column(0,0,25,self._cell_wrap_noborder)
            worksheet.set_column(1,50,7,self._cell_wrap_noborder)
             
            current_row += 2

            worksheet.write(
                current_row, 
                current_col, 
                "Employee Non Billable Hours", 
                self._main_header_format
            )       
            worksheet.write(current_row, current_col+1, "", self._main_header_format) 
            worksheet.write(current_row, current_col+2, "", self._main_header_format) 

            current_row += 1

            format_vertical_text = self._sub_header_format
            format_vertical_text.set_rotation(90)
            format_vertical_text.set_align('bottom')
            format_vertical_text.set_align('center')
            worksheet.set_row(current_row, 152)     

            employee_dict = {}
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
            self._start_employee_data_row = current_row
            
            for index,row in project_df.iterrows():
                worksheet.write(current_row, 0, row['Project']) 
                project_dict[row['Project'].strip()]=current_row
                current_row += 1
            
            self._end_employee_data_row = current_row
            
            for col in range(0,len(employee_dict)+1):
                worksheet.write(current_row, 
                    current_col+col, 
                    "", 
                    self._header_format_gray
                ) 
            current_row += 1
            group_per_project_emp = data_frame.groupby(['Project','Employee Name']) 
            for key, values in group_per_project_emp['Hours'].sum().items():
                worksheet.write(
                    project_dict[key[0].strip()], 
                    employee_dict[key[1].strip()], 
                    values
                )        
            current_row += 1

            worksheet.write(
                current_row, 
                current_col, 
                "Weekly Non Billable Hours", 
                self._main_header_format
            )       
            worksheet.write(current_row, current_col+1, "", self._main_header_format) 
            worksheet.write(current_row, current_col+2, "", self._main_header_format) 
            current_row += 1
           
            week_dict={}
            group_by_week = data_frame.groupby(['Week Range'])             
            week_df = pd.DataFrame(group_by_week.size().reset_index())
            worksheet.set_row(current_row, 100)    
            worksheet.write(current_row, 0, "", self._sub_header_format2)
            for index,row in week_df.iterrows():
                range_date = datetime.strptime(
                    str(row['Week Range']).strip(),
                    "%Y%m%d"
                )
                range_date_string ='{:02d}'.format(range_date.month) \
                    +'{:02d}'.format(range_date.day)
                key_string = f'Week Ending {range_date_string}'
                worksheet.write(
                    current_row, 
                    current_col+index+1, 
                    key_string, 
                    format_vertical_text
                ) 
                week_dict[row['Week Range'].strip()]=1+index
            current_row += 1

            self._start_weekl_data_row = current_row
            project_set2_dict = {}
            for index,row in project_df.iterrows():
                worksheet.write(current_row, 0, row['Project']) 
                project_set2_dict[row['Project'].strip()]=current_row
                current_row += 1
            self._end_weekl_data_row = current_row
            
            group_by_prj_week = data_frame.groupby(['Project','Week Range'])    
            for key, values in group_by_prj_week['Hours'].sum().items():
                worksheet.write(
                    project_set2_dict[key[0].strip()], 
                    week_dict[key[1].strip()], 
                    values
                )        
          
            for col in range(0,len(week_dict)+1):
                worksheet.write(current_row, current_col+col, "", self._header_format_gray) 
            current_row += 2

            worksheet.write(
                current_row, 
                current_col, 
                "Total Weekly Non Billable Hours", 
                self._main_header_format
            )       
            worksheet.write(current_row, current_col+1, "", self._main_header_format) 
            current_row += 1

            worksheet.set_row(current_row, 50)    
            worksheet.write(current_row, 0, "", self._sub_header_format2)
            worksheet.write(current_row, 1, 'Hours', format_vertical_text)
            
            current_row += 1

            grouped_per_proj_week =  data_frame.groupby(['Week Range'])
            self._start_week_total_row= current_row
            for key, values in grouped_per_proj_week['Hours'].sum().items():
                range_date = datetime.strptime(str(key).strip(),"%Y%m%d")
                range_date_string ='{:02d}'.format(range_date.month)+'{:02d}'.format(range_date.day)
                key_string = f'Week Ending {range_date_string}'
                worksheet.write(current_row, current_col, key_string)  
                worksheet.write(current_row, current_col+1, values, self._number_two_places)  
                current_row +=1
            self._end_week_total_row = current_row
            worksheet.write(current_row, current_col, "", self._header_format_gray)
            worksheet.write(current_row, current_col+1, "", self._header_format_gray)

            current_row += 2
            worksheet.write(
                current_row, 
                current_col, 
                "Total Non Billable Hours", 
                self._main_header_format
            )       
            worksheet.write(current_row, current_col+1, "", self._main_header_format) 
            current_row += 1
            worksheet.set_row(current_row, 45)
            worksheet.write(current_row, current_col, '', self._sub_header_format2)
            worksheet.write(current_row, current_col+1, 'Hours', format_vertical_text) 
            current_row +=1

            grouped_per_proj =  data_frame.groupby('Project')
            self._start_total_row= current_row
            for key, values in grouped_per_proj['Hours'].sum().items():
                worksheet.write(current_row, current_col, key)  
                worksheet.write(
                    current_row, 
                    current_col+1, 
                    values, 
                    self._number_two_places
                )  
                current_row +=1
            self._end_total_row = current_row
            worksheet.write(current_row, current_col, "", self._header_format_gray)
            worksheet.write(current_row, current_col+1, "", self._header_format_gray)
            
            self._create_charts(
                employee_dict, 
                project_dict, 
                week_dict, 
                project_set2_dict
            )

            end_row = 100
            end_row += 2
            worksheet.write(
                    end_row, 
                    0, 
                    f'Report generated : {datetime.now()} by {get_version()}', 
                    self._footer_format
                )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_charts(self, emp_dict, proj_dict, week_dict, proj_set2_dict):
        try:
            sheet_name = 'Charts'
                
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(13,1)
            )
            worksheet.merge_range(title_range,'','')


            worksheet.write(
                current_row, 
                current_col, 
                "Non Billable Hours Report Summary", 
                self._report_title
            )

            worksheet.set_column(0,0,25,self._cell_wrap_noborder)
            worksheet.set_column(1,50,8,self._cell_wrap_noborder)
            
          
            current_row += 2

            self._create_totals_chart(worksheet)
            self._create_weekly_totals_chart(worksheet,week_dict,proj_dict)
            self._create_weekly_chart(worksheet,week_dict,proj_set2_dict)
            self._create_employee_chart(worksheet, emp_dict,proj_dict)

        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_totals_chart(self, worksheet):
        try:
            chart = self._writer.book.add_chart({'type':'pie'}) 
                
            categories_address = get_cell_range_address(
                get_cell_address(0,self._start_total_row+1),
                get_cell_address(0,self._end_total_row +1),
                sheet='Data'
            )

            values_address = get_cell_range_address(
                get_cell_address(1, self._start_total_row+1),
                get_cell_address(1, self._end_total_row),
                sheet='Data'
            )
        
            chart.add_series({
                'categories': f'={categories_address}',
                'values':     f'={values_address}',
                'data_labels' : {'value':True}
            })

            
            chart.set_x_axis({'name': 'Employee'}) 
            chart.set_y_axis({'name': 'Hours'}) 
            chart.set_style(10)
            chart.set_title({'name':'Non Billable Hours Summary'})
            chart.set_legend({'position': 'bottom'})
            worksheet.insert_chart(
                'A2',
                chart,
                {'x_scale': 2, 'y_scale': 1.5,'x_offset':10,'y_offset':10}
            )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_weekly_totals_chart(self, worksheet, week_dict, project_dict):
        try:
            chart = self._writer.book.add_chart({'type':'line'}) 

            categories_address = get_cell_range_address(
                get_cell_address(0,self._start_week_total_row+1),
                get_cell_address(0,self._end_week_total_row+1),
                sheet='Data'
            )

            values_address = get_cell_range_address(
                get_cell_address(1,self._start_week_total_row+1),
                get_cell_address(1,self._end_week_total_row+1),
                sheet='Data'
            )

            chart.add_series({
                'categories': f'={categories_address}',
                'values':     f'={values_address}',
                'data_labels' : {'value':True}
            })
          
            chart.set_x_axis({'name': 'Weeks'}) 
            chart.set_y_axis({'name': 'Hours'}) 
            chart.set_style(10)
            chart.set_title({'name':'Weekly Total Non Billables Hours Summary'})
            chart.set_legend({'none': True})
            worksheet.insert_chart(
                'A24',
                chart,
                {'x_scale': 2, 'y_scale': 1,'x_offset':10,'y_offset':10}
            )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_weekly_chart(self, worksheet, week_dict, project_dict):
        try:
            column_chart = self._writer.book.add_chart({'type':'column'}) 
                
            categories_address = get_cell_range_address(
                get_cell_address(1,self._start_weekl_data_row),
                get_cell_address(len(week_dict),self._start_weekl_data_row),
                sheet='Data'
            )
            for key,value in project_dict.items():
                values_address = get_cell_range_address(
                    get_cell_address(1, value +1),
                    get_cell_address(len(week_dict),value+1),
                    sheet='Data'
                )
                column_chart.add_series({
                    'name': f'=Data!{get_cell_address(0,value+1)}',
                    'categories': f'={categories_address}',
                    'values':     f'={values_address}',
                    'data_labels' : {'value':True}
                })

            column_chart.set_y_axis({'name': 'Hours'}) 
            column_chart.set_style(10)
            column_chart.set_title({'name':'Weekly Non Billables Hours Summary'})
            column_chart.set_legend({'position': 'bottom'})
            worksheet.insert_chart(
                'A39',
                column_chart,
                {'x_scale': 2, 'y_scale': 1,'x_offset':10,'y_offset':10}
            )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_employee_chart(self, worksheet, employee_dict, project_dict):
        try:
            chart = self._writer.book.add_chart({'type':'column', 'subtype':'stacked'}) 
            
            categories_address = get_cell_range_address(
                get_cell_address(1, self._start_employee_data_row),
                get_cell_address(len(employee_dict),self._start_employee_data_row),
                sheet='Data'
            )
        
            for key,value in project_dict.items():
                values_address = get_cell_range_address(
                    get_cell_address(1,value+1),
                    get_cell_address(len(employee_dict),value+1),
                    sheet='Data'
                )
                
                chart.add_series({
                    'name': f'=Data!{get_cell_address(0,value+1)}',
                    'categories': f'={categories_address}',
                    'values':     f'={values_address}',
                    'data_labels' : {'value':True}
                })

           
            chart.set_x_axis({'name': 'Employee'}) 
            chart.set_y_axis({'name': 'Hours'}) 
            chart.set_style(10)
            chart.set_title({'name':'Employee Non Billable Hours Summary'})
            worksheet.insert_chart(
                'A54',
                chart,
                {'x_scale': 2, 'y_scale': 2,'x_offset':10,'y_offset':10}
            )
 
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)



  