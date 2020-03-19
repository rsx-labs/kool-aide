import json
import pandas as pd
from datetime import datetime
import xlsxwriter
from typing import List

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import get_cell_address, get_cell_range_address

from kool_aide.model.cli_argument import CliArgument
from kool_aide.assets.resources.messages import *

class AssetInventoryReport:

    def __init__(self, logger: CustomLogger, settings: AppSetting, 
                    data: pd.DataFrame, writer = None) -> None:
        
        self._data = data
        self._settings = settings,
        self._logger = logger

        self._writer = writer
        self._workbook = None
        self._main_header_format = None
        self._report_title = None
        self._header_format_orange = None
        self._header_format_gray = None
        self._cell_wrap_noborder = None
        self._cell_total = None
        self._cell_sub_total = None

        self._log('creating component ...')

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report.asset_inventory_report]", level)

    def generate(self, format: str) -> None:
        if format == OUTPUT_FORMAT[3]: # excel
            return self._generate_excel_report()
        else:
            return False, NOT_SUPPORTED

    def _generate_excel_report(self) -> None:
        try:
            data_frame = self._data
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            wrap_content = self._workbook.add_format(SHEET_CELL_WRAP)
            report_title = self._workbook.add_format(SHEET_TITLE)

            drop_columns=[
                'DateAssigned',
                'DatePurchased',
                'DepartmentID',
                'DivisionID',
                'StatusID'
            ]

            column_headers = [
                'Employee ID',
                'Employee Name',
                'Department',
                'Description',
                'Manufacturer',
                'Model',
                'Serial Number',
                'Asset Tag',
                'Other Info',
                'Comments',
                'Status'
            ]
            
            data_frame.drop(drop_columns, inplace=True, axis=1)
            data_frame.columns = column_headers
        
           
            data_frame.to_excel(
                self._writer, 
                sheet_name='Asset Inventory', 
                index= False ,
                startrow = 1
            )

            worksheet = self._writer.sheets['Asset Inventory']
            worksheet.set_column(0,0,12)
            worksheet.set_column(1,1,35)
            worksheet.set_column(2,3,14)
            worksheet.set_column(4,7,22)
            worksheet.set_column(8,9,25, wrap_content)
            worksheet.set_column(10,10,15)
            current_row = 0
            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(len(column_headers) - 1,1)
            )
            worksheet.merge_range(title_range,'','')
            worksheet.write(current_row,0,'Asset Inventory Report', report_title)
            current_row += 1
            for col_num, value in enumerate(data_frame.columns.values):
                worksheet.write(current_row, col_num, value, main_header_format)       

            total_row = len(data_frame)
            worksheet.write(
                total_row + 4, 
                0, 
                f'Report generated : {datetime.now()}', 
                footer_format
            )
               
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            # print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)