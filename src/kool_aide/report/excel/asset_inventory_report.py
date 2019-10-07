import json
import pandas as pd
from datetime import datetime
import xlsxwriter
from typing import List

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *

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

            drop_columns=[
                'DateAssigned',
                'DatePurchased',
                'DepartmentID',
                'DivisionID'
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
                'Comments'
            ]
            
            data_frame.drop(drop_columns, inplace=True, axis=1)
            data_frame.columns = column_headers
        
            data_frame.to_excel(
                self._writer, 
                sheet_name='Asset Inventory', 
                index= False 
            )

            worksheet = self._writer.sheets['Asset Inventory']
            worksheet.set_column(0,0,12)
            worksheet.set_column(1,1,35)
            worksheet.set_column(2,3,14)
            worksheet.set_column(4,7,22)
            worksheet.set_column(8,9,25, wrap_content)
            
            for col_num, value in enumerate(data_frame.columns.values):
                worksheet.write(0, col_num, value, main_header_format)       

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