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

class KPISummaryReport:

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
        self._month_dict = {}
        self._subject_dict = {}

        self._log('creating component ...')

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report.kpi_summary_report]", level)

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
            self._header_format_lt_gray = self._workbook.add_format(SHEET_HEADER_LT_GRAY)

            self._create_data()
            self._create_summary()
               
            self._writer.save()
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_summary(self):
        try:
            data_frame = self._data
            sheet_name = 'KPISummary'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            worksheet.set_column(0,0,25,self._cell_wrap_noborder)
            worksheet.set_column(1,50,7,self._cell_wrap_noborder)
       
            current_row = 21
     
            group_by_month = data_frame.groupby(['MonthID','Month'])             
            by_month_df = pd.DataFrame(group_by_month.size().reset_index())
            
            month2_dict = {}
            start_row = current_row

            header = self._main_header_format
            header.set_align('center')
            sub_header = self._sub_header_format
            sub_header.set_align('center')
            sub_header2 = self._sub_header_format2
            sub_header2.set_align('center')

            worksheet.write(current_row, current_col, "Month", header)
            current_col = 1
            for index,row in by_month_df.iterrows():
                cell_range = get_cell_range_address(
                    get_cell_address(current_col, current_row + 1),
                    get_cell_address(current_col + 1, current_row + 1)
                )
                worksheet.merge_range(cell_range,'','')
                worksheet.write(
                    current_row, 
                    current_col, 
                    str(row['Month']).strip(),
                    sub_header2
                )
                worksheet.write(
                    current_row+1, 
                    current_col, 
                    'Target',
                    sub_header
                )
                worksheet.write(
                    current_row+1, 
                    current_col+1, 
                    'Actual',
                    sub_header
                )

                month2_dict[row['Month'].strip()]=current_col
                current_col += 2

            current_col = 0
            worksheet.write(current_row+1, current_col, '', self._sub_header_format2)
            current_row +=2

            header_len = len(month2_dict) * 2

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(header_len,1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(
                0, 
                0, 
                "KPI Target Summary Report", 
                self._report_title
            )

            group_by_subject = data_frame.groupby(['Subject'])             
            subject_df = pd.DataFrame(group_by_subject.size().reset_index())
           
            subject2_dict = {}
            for index,row in subject_df.iterrows():
                worksheet.write(current_row,0,row['Subject'])   
                subject2_dict[row['Subject'].strip()]=current_row
                current_row += 1
            
            last_row = current_row

            for index,row in data_frame.iterrows():
                worksheet.write(
                    subject2_dict[row['Subject'].strip()],
                    month2_dict[row['Month'].strip()],                  
                    float(row['Target']) * 100
                )  
                worksheet.write(
                    subject2_dict[row['Subject'].strip()],
                    month2_dict[row['Month'].strip()]+1,           
                    float(row['Actual']) * 100
                )  
                #current_row += 1
            
            worksheet.write(current_row, 0, 'Overall', self._header_format_lt_gray)

            group_by_subject_month = data_frame.groupby(['Month'])
            overall_df = group_by_subject_month['Overall'].sum().reset_index()
            for key, row in overall_df.iterrows():
                col_temp = month2_dict[row['Month']]
                cell_range = get_cell_range_address(
                    get_cell_address(col_temp, current_row + 1),
                    get_cell_address(col_temp + 1, current_row + 1)
                )
                worksheet.merge_range(cell_range,'','')

                worksheet.write(
                    current_row, 
                    month2_dict[row['Month']], 
                    (float(row['Overall'])/3) * 100
                )
            
            current_row +=1

            try:
                chart = self._writer.book.add_chart({'type':'column'}) 
                
                categories_address = get_cell_range_address(
                    get_cell_address(1,4),
                    get_cell_address(len(self._month_dict),4),
                    sheet='Data'
                )

                last_charted_row = 0
                for key,value in self._subject_dict.items():
                    last_charted_row = value + 1
                    values_address = get_cell_range_address(
                        get_cell_address(1,value+1),
                        get_cell_address(len(self._month_dict),value+1),
                        sheet='Data'
                    )
                    chart.add_series({
                        'name': f'=Data!{get_cell_address(0,value+1)}',
                        'categories': f'={categories_address}',
                        'values':     f'={values_address}',
                        'data_labels' : {'value':False}
                    })

            
                chart2 = self._writer.book.add_chart({'type':'line'}) 
                
                categories_address = get_cell_range_address(
                    get_cell_address(1,4),
                    get_cell_address(len(self._month_dict),4),
                    sheet='Data'
                )
                
                values_address = get_cell_range_address(
                    get_cell_address(1,last_charted_row+1),
                    get_cell_address(len(self._month_dict),last_charted_row + 1),
                    sheet='Data'
                )
                chart2.add_series({
                    'name': f'=Data!{get_cell_address(0,last_charted_row+1)}',
                    'categories': f'={categories_address}',
                    'values':     f'={values_address}',
                    'data_labels' : {'value':False}
                })
                chart.combine(chart2)
            except:
                pass
            chart.set_x_axis({'name': 'Month'}) 
            chart.set_y_axis({'name': 'KPI %'}) 
            chart.set_style(10)
            chart.set_title({'name':'KPI Summary'})
            worksheet.insert_chart(
                'A2',
                chart,
                {'x_scale': 2.12, 'y_scale': 1.3,'x_offset':10,'y_offset':10}
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

    def _create_data(self):
        try:
            data_frame = self._data
            sheet_name = 'Data'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            worksheet.set_column(0,0,25,self._cell_wrap_noborder)
            worksheet.set_column(1,50,11,self._cell_wrap_noborder)
       
            current_row = 2
     
            group_by_month = data_frame.groupby(['MonthID','Month'])             
            by_month_df = pd.DataFrame(group_by_month.size().reset_index())
            
            self._month_dict = {}
            start_row = current_row

            header = self._sub_header_format
            header.set_align('center')
            sub_header = self._header_format_lt_gray
            sub_header.set_align('center')

            worksheet.write(current_row, current_col, 'Actual KPI', self._main_header_format)
            current_row +=1

            worksheet.write(current_row, current_col, "Month", header)
            current_col = 1
            for index,row in by_month_df.iterrows():
                worksheet.write(
                    current_row, 
                    current_col + index, 
                    str(row['Month']).strip(),
                    header
                )
                
                self._month_dict[row['Month'].strip()]=index + 1
            current_row += 1
            current_col = 0
          
            header_len = len(self._month_dict) 

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(header_len,1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(
                0, 
                0, 
                "KPI Target Summary Report", 
                self._report_title
            )

            group_by_subject = data_frame.groupby(['Subject'])             
            subject_df = pd.DataFrame(group_by_subject.size().reset_index())
           
            self._subject_dict = {}
            for index,row in subject_df.iterrows():
                worksheet.write(current_row,0,row['Subject'])   
                self._subject_dict[row['Subject'].strip()]=current_row
                current_row += 1
            
            last_row = current_row

            for index,row in data_frame.iterrows():
                worksheet.write(
                    self._subject_dict[row['Subject'].strip()],
                    self._month_dict[row['Month'].strip()],           
                    float(row['Actual']) * 100
                )  
                #current_row += 1
            
            worksheet.write(current_row, 0, 'Overall', self._header_format_lt_gray)

            group_by_subject_month = data_frame.groupby(['Month'])
            overall_df = group_by_subject_month['Overall'].sum().reset_index()
            for key, row in overall_df.iterrows():
                
                worksheet.write(
                    current_row, 
                    self._month_dict[row['Month']], 
                    (float(row['Overall'])/3) * 100
                )
            
            current_row +=2

            worksheet.write(current_row, current_col, 'Target KPI', self._main_header_format)
            current_row +=1

            worksheet.write(current_row, current_col, "Month", header)
            current_col = 1
            for index,row in by_month_df.iterrows():
                worksheet.write(
                    current_row, 
                    current_col + index, 
                    str(row['Month']).strip(),
                    header
                )
                
                self._month_dict[row['Month'].strip()]=index + 1
            current_row += 1
            current_col = 0

            subject2_dict ={}
            for index,row in subject_df.iterrows():
                worksheet.write(current_row,0,row['Subject'])   
                subject2_dict[row['Subject'].strip()]=current_row
                current_row += 1
            
        
            for index,row in data_frame.iterrows():
                worksheet.write(
                    subject2_dict[row['Subject'].strip()],
                    self._month_dict[row['Month'].strip()],           
                    float(row['Target']) * 100
                )  
  

            end_row = current_row
            end_row += 2
            worksheet.write(
                end_row, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                self._footer_format
            )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    