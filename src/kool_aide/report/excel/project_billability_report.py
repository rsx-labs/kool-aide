import json
import pandas as pd
from datetime import datetime
import xlsxwriter
from typing import List

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import get_cell_address, get_cell_range_address, get_version

from kool_aide.model.cli_argument import CliArgument
from kool_aide.assets.resources.messages import *

class ProjectBillabilityReport:

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
        self._logger.log(f"{message} [report.project_billability_report]", level)

    def generate(self, format: str) -> None:
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
            self._header_format_orange = self._workbook.add_format(SHEET_HEADER_ORANGE)
            self._header_format_gray = self._workbook.add_format(SHEET_HEADER_GRAY)
            self._number_two_places = self._workbook.add_format({'num_format':'0.00'})
        
            self._create_month_summary()
            self._create_week_summary()
               
            self._writer.save()
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_month_summary(self):
        try:
            data_frame = self._data
            grouped_per_project =  data_frame.groupby('Project')
            sheet_name = 'Summary'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)
            worksheet.write(
                current_row, 
                current_col, 
                "Billability Report", 
                self._main_header_format
            )       
            worksheet.set_column(0,0,40,self._cell_wrap_noborder)
            worksheet.set_column(1,1,15,self._cell_wrap_noborder)
            worksheet.set_column(2,2,15,self._cell_wrap_noborder)
            worksheet.set_column(3,3,5,self._cell_wrap_noborder)
            worksheet.set_column(4,4,40,self._cell_wrap_noborder)
            worksheet.set_column(5,5,15,self._cell_wrap_noborder)
            worksheet.set_column(6,6,15,self._cell_wrap_noborder)

            for col in range(1,7):
                worksheet.write(
                    current_row, 
                    current_col+col, 
                    "", 
                    self._main_header_format
                )
                 
            current_row = 28
     
            worksheet.write(
                current_row, 
                current_col, 
                "Billability Per Project", 
                self._header_format_orange
            )

            worksheet.write(current_row, current_col+1, "", self._header_format_orange) 
            worksheet.write(current_row, current_col+2, "", self._header_format_orange) 

            current_row += 1

            worksheet.write(current_row, current_col, "Project", self._header_format_gray)       
            worksheet.write(current_row, current_col+1, "Hours", self._header_format_gray) 
            worksheet.write(current_row, current_col+2, "Percentage", self._header_format_gray) 

            current_row += 1
            start_row = current_row
            total = sum(data_frame['Hours'])
            
            df_sum_project_hours = grouped_per_project['Hours'].sum().reset_index()
            df_sum_project_hours.sort_values(by=['Hours'], inplace=True)
            
            for key, row in df_sum_project_hours.iterrows():
                worksheet.write(current_row, current_col, row['Project'])       
                worksheet.write(current_row, current_col+1, row['Hours'], self._number_two_places)
                percentage = (float(row['Hours']) / total) * 100
                worksheet.write(current_row, current_col+2, percentage, self._number_two_places) 
                current_row +=1

            worksheet.write(current_row, current_col, "Total Hours", self._header_format_gray) 
            worksheet.write(current_row, current_col+1, total, self._header_format_gray) 
            worksheet.write(current_row, current_col+2, 100.00, self._header_format_gray) 

            end_row=current_row
            categories_address = get_cell_range_address(
                get_cell_address(0,start_row + 1),
                get_cell_address(0,current_row),
                sheet=sheet_name
            )

            values_address = get_cell_range_address(
                get_cell_address(1, start_row + 1),
                get_cell_address(1, current_row),
                sheet=sheet_name
            )

            chart = self._writer.book.add_chart({'type':'bar'})
            chart.add_series({
                'categories': f'={categories_address}',
                'values':     f'={values_address}',
                'data_labels' : {'value':True}
                })
            
            chart.set_x_axis({'name': 'Hours'}) 
            chart.set_y_axis({'name': 'Projects'}) 
            chart.set_legend({'none': True})
            chart.set_style(20)
            chart.set_title({'name':'Project Billability Summary'})
            
            worksheet.insert_chart(
                'A2',
                chart,
                {'x_scale': 1, 'y_scale': 1.7,'x_offset':10,'y_offset':10}
            )

            current_row = 28
            current_col = 4
     
            worksheet.write(
                current_row, 
                current_col, 
                "Billable \ Non Billable", 
                self._header_format_orange
            )   

            worksheet.write(current_row, current_col+1, "", self._header_format_orange) 
            worksheet.write(current_row, current_col+2, "", self._header_format_orange) 
            current_row += 1

            worksheet.write(current_row, current_col, "Project", self._header_format_gray)       
            worksheet.write(current_row, current_col+1, "Hours", self._header_format_gray)  
            worksheet.write(current_row, current_col+2, "Percentage", self._header_format_gray) 
            current_row += 1

            grouped_per_billability =  data_frame.groupby('IsBillable')

            start_row = current_row
            for key, values in grouped_per_billability['Hours'].sum().items():
                descr = 'Billable' if key==1 else 'Non-billable'   
                worksheet.write(current_row, current_col, descr)  
                worksheet.write(current_row, current_col+1, values, self._number_two_places)  
                percentage = (float(values) / total) * 100
                worksheet.write(current_row, current_col+2, percentage, self._number_two_places) 
                current_row +=1

            worksheet.write(current_row, current_col, "Total Hours", self._header_format_gray) 
            worksheet.write(current_row, current_col+1, total, self._header_format_gray) 
            worksheet.write(current_row, current_col+2, 100.00, self._header_format_gray) 

            categories_address = get_cell_range_address(
                get_cell_address(4,start_row + 1),
                get_cell_address(4,current_row),
                sheet=sheet_name
            )

            values_address = get_cell_range_address(
                get_cell_address(5, start_row + 1),
                get_cell_address(5, current_row),
                sheet=sheet_name
            )

            chart = self._writer.book.add_chart({'type':'pie'})
            chart.add_series({
                'categories': f'={categories_address}',
                'values':     f'={values_address}',
                'data_labels' : {'value':True}
                })
            
            chart.set_style(10)
            chart.set_title({'name':'Billable vs Non-Billable'})
            chart.set_legend({'position': 'top'})
            worksheet.insert_chart(
                'E2',
                chart,
                {'x_scale': 1, 'y_scale': 1.7,'x_offset':10,'y_offset':10}
            )

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
            # data_frame.sort_values(by=['Month'], inplace= True)
            grouped_per_week =  data_frame.groupby(['Week Range'])
            for key, values in grouped_per_week:
                df_per_week = pd.DataFrame(grouped_per_week.get_group(key))
            
                range_date = datetime.strptime(key,"%Y%m%d")
                range_date_string ='{:02d}'.format(range_date.month)+'{:02d}'.format(range_date.day)
                key_string = f'WeekEnding_{range_date_string}'
                week_ending = f'Week Ending {range_date_string}'
                sheet_name = key_string
            
                current_row = 0
                current_col = 0

                worksheet = self._writer.book.add_worksheet(sheet_name)
                worksheet.write(
                    current_row, 
                    current_col, 
                    f"Billability Report for {week_ending", 
                    self._main_header_format
                )

                worksheet.set_column(0,0,40,self._cell_wrap_noborder)
                worksheet.set_column(1,1,15,self._cell_wrap_noborder)
                worksheet.set_column(2,2,15,self._cell_wrap_noborder)
                worksheet.set_column(3,3,5,self._cell_wrap_noborder)
                worksheet.set_column(4,4,40,self._cell_wrap_noborder)
                worksheet.set_column(5,5,15,self._cell_wrap_noborder)
                worksheet.set_column(6,6,15,self._cell_wrap_noborder)

                for col in range(1,7):
                    worksheet.write(current_row, current_col+col, "", self._main_header_format) 
                current_row += 2

                current_row = 28
        
                worksheet.write(
                    current_row, 
                    current_col, 
                    "Billability Per Project", 
                    self._header_format_orange
                )

                worksheet.write(current_row, current_col+1, "", self._header_format_orange) 
                worksheet.write(current_row, current_col+2, "", self._header_format_orange) 

                current_row += 1

                worksheet.write(current_row, current_col, "Project", self._header_format_gray)       
                worksheet.write(current_row, current_col+1, "Hours", self._header_format_gray) 
                worksheet.write(current_row, current_col+2, "Percentage", self._header_format_gray) 

                current_row += 1
                start_row = current_row

                grouped_per_project = df_per_week.groupby(['Project'])
                total = sum(df_per_week['Hours'])

                df_sum_project_hours = grouped_per_project['Hours'].sum().reset_index()
                df_sum_project_hours.sort_values(by=['Hours'], inplace=True)
                for key, row in df_sum_project_hours.iterrows():
                    worksheet.write(current_row, current_col, row['Project'])       
                    worksheet.write(current_row, current_col+1, row['Hours'], self._number_two_places)
                    percentage = (float(row['Hours']) / total) * 100
                    worksheet.write(current_row, current_col+2, percentage, self._number_two_places) 
                    current_row +=1

                worksheet.write(current_row, current_col, "Total Hours", self._header_format_gray) 
                worksheet.write(current_row, current_col+1, total, self._header_format_gray) 
                worksheet.write(current_row, current_col+2, 100.00, self._header_format_gray) 

                end_row=current_row
                categories_address = get_cell_range_address(
                    get_cell_address(0,start_row + 1),
                    get_cell_address(0,current_row),
                    sheet=sheet_name
                )

                values_address = get_cell_range_address(
                    get_cell_address(1, start_row + 1),
                    get_cell_address(1, current_row),
                    sheet=sheet_name
                )

                chart = self._writer.book.add_chart({'type':'bar'})
                chart.add_series({
                    'categories': f'={categories_address}',
                    'values':     f'={values_address}',
                    'data_labels' : {'value':True}
                    })
                
                chart.set_x_axis({'name': 'Hours'}) 
                chart.set_y_axis({'name': 'Projects'}) 
                chart.set_legend({'none': True})
                chart.set_style(25)
                chart.set_title({'name':'Project Billability for the Week'})
                
                worksheet.insert_chart(
                    'A2',
                    chart,
                    {'x_scale': 1, 'y_scale': 1.7,'x_offset':10,'y_offset':10}
                )

                current_row = 28
                current_col = 4
        
                worksheet.write(
                    current_row, 
                    current_col, 
                    "Billable \ Non Billable", 
                    self._header_format_orange
                )

                worksheet.write(current_row, current_col+1, "", self._header_format_orange) 
                worksheet.write(current_row, current_col+2, "", self._header_format_orange) 
                current_row += 1

                worksheet.write(current_row, current_col, "Project", self._header_format_gray)       
                worksheet.write(current_row, current_col+1, "Hours", self._header_format_gray)  
                worksheet.write(current_row, current_col+2, "Percentage", self._header_format_gray) 
                current_row += 1

                grouped_per_billability =  df_per_week.groupby('IsBillable')

                start_row = current_row
                for key, values in grouped_per_billability['Hours'].sum().items():
                    descr = 'Billable' if key==1 else 'Non-billable'   
                    worksheet.write(current_row, current_col, descr)  
                    worksheet.write(current_row, current_col+1, values, self._number_two_places)  
                    percentage = (float(values) / total) * 100
                    worksheet.write(current_row, current_col+2, percentage, self._number_two_places) 
                    current_row +=1

                worksheet.write(current_row, current_col, "Total Hours", self._header_format_gray) 
                worksheet.write(current_row, current_col+1, total, self._header_format_gray) 
                worksheet.write(current_row, current_col+2, 100.00, self._header_format_gray) 

                categories_address = get_cell_range_address(
                    get_cell_address(4,start_row + 1),
                    get_cell_address(4,current_row),
                    sheet=sheet_name
                )

                values_address = get_cell_range_address(
                    get_cell_address(5, start_row + 1),
                    get_cell_address(5, current_row),
                    sheet=sheet_name
                )

                chart = self._writer.book.add_chart({'type':'pie'})
                chart.add_series({
                    'categories': f'={categories_address}',
                    'values':     f'={values_address}',
                    'data_labels' : {'value':True}
                    })
                
                chart.set_style(10)
                chart.set_title({'name':'Billable vs Non-Billable'})
                chart.set_legend({'position': 'top'})
                
                worksheet.insert_chart(
                    'E2',
                    chart,
                    {'x_scale': 1, 'y_scale': 1.7,'x_offset':10,'y_offset':10}
                )

                end_row += 2
                worksheet.write(
                        end_row, 
                        0, 
                        f'Report generated : {datetime.now()} by {get_version()}', 
                        self._footer_format
                    )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)



    