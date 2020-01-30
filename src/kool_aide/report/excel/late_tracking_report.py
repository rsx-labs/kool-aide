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

class LateTrackingReport:

    def __init__(self, logger: CustomLogger, settings: AppSetting, 
                    data: pd.DataFrame, writer = None) -> None:
        
        self._data = data[data["StatusID"] == 11]
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

        self._month1_dict ={}
        self._month2_dict = {}
        self._employee_dict = {}

        self._log('creating component ...')

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report.late_tracking_report]", level)

    def generate(self, format: str) -> None:
        if format == OUTPUT_FORMAT[3]: # excel
            return self._generate_excel_report()
        else:
            return False, NOT_SUPPORTED

    def _generate_excel_report(self) -> None:
        try:
            self._workbook = self._writer.book
            self._main_header_format = self._workbook.add_format(SHEET_TOP_HEADER_NO_WRAP)
            self._footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            self._wrap_content = self._workbook.add_format(SHEET_CELL_WRAP)
            self._sub_header = self._workbook.add_format(SHEET_SUB_HEADER)
            self._sub_header2 = self._workbook.add_format(SHEET_SUB_HEADER2)
            self._header_format_gray = self._workbook.add_format(SHEET_HEADER_GRAY)
            self._number_two_places = self._workbook.add_format({'num_format':'0.00'})
            self._report_title = self._workbook.add_format(SHEET_TITLE)
            self._total = self._workbook.add_format(SHEET_HEADER_LT_GRAY)
            
            self._total.set_align('center')
            self._total.set_align('vcenter')
            
            self._create_by_month()
            self._create_by_employee()
               
            self._writer.save()
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

   
    def _create_by_month(self):
        try:
            data_frame = self._data
            sheet_name = 'By_Month'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(12,1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(
                current_row, 
                current_col, 
                'Late Tracking Report', 
                self._report_title
            )

            worksheet.set_column(0,0,25,self._cell_wrap_noborder)
            worksheet.set_column(1,50,5,self._cell_wrap_noborder)
            worksheet.set_row(17, 70)  
            worksheet.set_row(35, 70)     

            current_row += 16
     
            worksheet.write(
                current_row, 
                current_col, 
                "Monthly Late Summary", 
                self._main_header_format
            )

            current_row += 1

            group_by_month = data_frame.groupby(['MonthID','Month'])             
            month_df = pd.DataFrame(group_by_month.size().reset_index())
            format_vertical_text = self._sub_header
            format_vertical_text.set_rotation(90)
            format_vertical_text.set_align('bottom')
            start_row = current_row

            worksheet.write(current_row, current_col, "", self._sub_header2)
            month_idx = 0
            for month in FISCAL_MONTHS:
                worksheet.write(
                    current_row, 
                    current_col+month_idx+1, 
                    month, 
                    format_vertical_text
                )

                self._month1_dict[month]=month_idx
                month_idx += 1

            current_row +=1
            group_per_month = data_frame.groupby(['Month']) 
            worksheet.write(current_row, 0, "Total Count", self._total)
            for month in FISCAL_MONTHS:
                count = 0
                try:
                    count = len(pd.DataFrame(group_per_month.get_group(month)))
                except:
                    pass
                worksheet.write(
                    current_row,
                    self._month1_dict[month]+1, 
                    count
                )  
            
            last_row= current_row
            
            chart1 = self._writer.book.add_chart({'type':'line'}) 
                
            categories_address = get_cell_range_address(
                get_cell_address(1,18),
                get_cell_address(len(self._month1_dict),18),
                sheet_name
            )
            
            values_address = get_cell_range_address(
                get_cell_address(1,19),
                get_cell_address(len(self._month1_dict),19),
                sheet_name
            )
            
            chart1.add_series({
                'name': f'={sheet_name}!{get_cell_address(0,18)}',
                'categories': f'={categories_address}',
                'values':     f'={values_address}',
                'data_labels' : {'value':True}
            })

            chart1.set_x_axis({'name': 'Month'}) 
            chart1.set_y_axis({'name': 'No. of Lates'}) 
            chart1.set_style(10)
            chart1.set_title({'name':'Total Lates Per Month'})
            chart1.set_legend({'none': True})
            worksheet.insert_chart(
                'A2',
                chart1,
                {'x_scale':1.34, 'y_scale': 0.96,'x_offset':10,'y_offset':10}
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

    def _create_by_employee(self):
        try:
            data_frame = self._data
            sheet_name = 'By_Employee'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(24,1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(
                current_row, 
                current_col, 
                'Late Tracking Report', 
                self._report_title
            )

            worksheet.set_column(0,0,30,self._cell_wrap_noborder)
            worksheet.set_column(1,50,5,self._cell_wrap_noborder)
            worksheet.set_row(26, 70)  
           
            current_row += 25
     
            worksheet.write(
                current_row, 
                current_col, 
                "Employee Monthly Late Summary", 
                self._main_header_format
            )

            current_row += 1

            group_by_month = data_frame.groupby(['MonthID','Month'])             
            month_df = pd.DataFrame(group_by_month.size().reset_index())
            format_vertical_text = self._sub_header
            format_vertical_text.set_rotation(90)
            format_vertical_text.set_align('bottom')
            start_row = current_row

            worksheet.write(current_row, current_col, "", self._sub_header2)
            month_idx = 0
            for month in FISCAL_MONTHS:
                worksheet.write(
                    current_row, 
                    current_col+month_idx+1, 
                    month, 
                    format_vertical_text
                )

                self._month1_dict[month]=month_idx
                month_idx += 1
            
            worksheet.write(
                current_row, 
                len(self._month1_dict)+1, 
                "Total", 
                self._header_format_gray
            )

            current_row +=1
            
            group_by_emp = data_frame.groupby(['Employee Name'])             
            employee_df = pd.DataFrame(group_by_emp.size().reset_index())
            for index,row in employee_df.iterrows():
                worksheet.write(current_row, 0, row['Employee Name'].strip()) 
                self._employee_dict[row['Employee Name'].strip()]=current_row
                current_row += 1
            last_row = current_row
            
            group_per_month_emp = data_frame.groupby(['Month', 'Employee Name']) 
            month_employee_df = pd.DataFrame(group_per_month_emp.size().reset_index())
          
            count_per_employee_month = group_per_month_emp['Employee Name'].count()
            for key,values in count_per_employee_month.iteritems():
                worksheet.write(
                    self._employee_dict[key[1].strip()], 
                    self._month1_dict[key[0]]+1,
                    values
                )  

            count_per_employee = group_by_emp['Employee Name'].count()
            for key,values in count_per_employee.iteritems():
                worksheet.write(
                    self._employee_dict[key.strip()], 
                    len(self._month1_dict)+1,
                    values
                )  

        
            chart1 = self._writer.book.add_chart({'type':'column','subtype':'stacked'}) 
                
            categories_address = get_cell_range_address(
                get_cell_address(1,27),
                get_cell_address(len(self._month1_dict),27),
                sheet_name
            )

            for key,value in self._employee_dict.items():
                values_address = get_cell_range_address(
                    get_cell_address(1,value+1),
                    get_cell_address(len(self._month1_dict),value+1),
                    sheet=sheet_name
                )
                chart1.add_series({
                    'name': f'={sheet_name}!{get_cell_address(0,value+1)}',
                    'categories': f'={categories_address}',
                    'values':     f'={values_address}',
                    'data_labels' : {'value':True}
                })
            
       
            chart1.set_x_axis({'name': 'Month'}) 
            chart1.set_y_axis({'name': 'No. of Lates'}) 
            chart1.set_style(10)
            chart1.set_title({'name':'Monthly Employee Lates'})
            chart1.set_legend({'position': 'bottom'})
            worksheet.insert_chart(
                'A2',
                chart1,
                {'x_scale':1.5, 'y_scale': 1.6,'x_offset':10,'y_offset':10}
            )
            
            chart2 = self._writer.book.add_chart({'type':'bar'}) 
                
            categories_address = get_cell_range_address(
                get_cell_address(0,28),
                get_cell_address(0,28+ len(self._employee_dict)-1),
                sheet_name
            )
       
            values_address = get_cell_range_address(
                get_cell_address(13,28),
                get_cell_address(13,28 + len(self._employee_dict)-1),
                sheet=sheet_name
            )
            chart2.add_series({
                
                'categories': f'={categories_address}',
                'values':     f'={values_address}',
                'data_labels' : {'value':True}
            })
            
       
            chart2.set_x_axis({'name': 'Lates'}) 
            chart2.set_y_axis({'name': 'Employee'}) 
            chart2.set_style(10)
            chart2.set_title({'name':'Total Lates Per Employee'})
            chart2.set_legend({'position': 'none'})
            worksheet.insert_chart(
                'P2',
                chart2,
                {'x_scale':0.8, 'y_scale': 1.6,'x_offset':10,'y_offset':10}
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
