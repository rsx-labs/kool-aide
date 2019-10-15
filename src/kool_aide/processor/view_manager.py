# kool-aide/processor/common_manager.py

import pprint
import jsonpickle
import json
from beautifultable import BeautifulTable
import pandas as pd
from tabulate import tabulate
import os
from datetime import datetime
import xlsxwriter
from typing import List
from datetime import datetime


from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import append_date_to_file_name

from kool_aide.db_access.connection import Connection
from kool_aide.db_access.dbhelper.view_helper \
    import ViewHelper

from kool_aide.model.cli_argument import CliArgument
from kool_aide.model.aide.project import Project
from kool_aide.model.aide.week_range import WeekRange

from kool_aide.assets.resources.messages import *


class ViewManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument = None):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments

        self._db_helper = ViewHelper(
            self._logger, 
            self._config, 
            self._connection
        )

        self._report_settings = None
        self._report_schedules = None

        self._load_report_settings()

        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [processor.view_manager]", level)

    def retrieve(self, arguments : CliArgument):
        self._log(f"retrieving view : {arguments.model}")
  
        if arguments.view == SUPPORTED_VIEWS[0]:
            self._retrieve_status_report_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[1]:
            self._retrieve_asset_inventory_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[2]:
            self._retrieve_project_view(arguments)
            return True, DATA_RETRIEVED
        
        return False, NOT_SUPPORTED

    def get_status_report_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        project_filter = None
        week_filter = None
        division_filter = None
        department_filter = None

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    project_filter = None if PARAM_PROJECT not in json_parameters else json_parameters[PARAM_PROJECT] 
                    week_filter = None if PARAM_WEEK not in json_parameters else json_parameters[PARAM_WEEK] 
                    division_filter = None if PARAM_DIVISIONS not in json_parameters else json_parameters[PARAM_DIVISIONS] 
                    department_filter = None if PARAM_DEPARTMENTS not in json_parameters else json_parameters[PARAM_DEPARTMENTS] 
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            # if running as auto, use settings. else, use params
            if arguments.auto_mode:
                week_filter = self._get_report_schedule(datetime.now().month)
           
            results = self._db_helper.get_status_report_view(week_filter)
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if project_filter is not None and len(project_filter)>0:
                data_frame = data_frame[data_frame['Project'].isin(project_filter)]

            if division_filter is not None and len(division_filter)>0:
                data_frame = data_frame[data_frame['DivisionID'].isin(division_filter)]
            
            if department_filter is not None and len(department_filter)>0:
                data_frame = data_frame[data_frame['DepartmentID'].isin(department_filter)]

            if sort_keys is not None and len(sort_keys)>0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns)>0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
            
            data_frame = data_frame[data_frame['ActualWeekWork'] > 0]

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
            return None
    
    def get_asset_inventory_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        divisions = None
        departments = None
        
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    divisions = None if PARAM_DIVISIONS not in json_parameters else json_parameters[PARAM_DIVISIONS] 
                    departments = None if PARAM_DEPARTMENTS not in json_parameters else json_parameters[PARAM_DEPARTMENTS] 
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_asset_inventory_view()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if departments is not None and len(departments)>0:
                data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns) > 0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
            
            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_project_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        divisions = None
        departments = None
        projects = None
        
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    divisions = None if PARAM_DIVISIONS not in json_parameters else json_parameters[PARAM_DIVISIONS] 
                    departments = None if PARAM_DEPARTMENTS not in json_parameters else json_parameters[PARAM_DEPARTMENTS] 
                    projects = None if PARAM_PROJECT not in json_parameters else json_parameters[PARAM_PROJECT] 
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_project_view()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if projects is not None and len(projects)>0:
                data_frame = data_frame[data_frame['ProjectName'].isin(projects)]

            if departments is not None and len(departments)>0:
                data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns) > 0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
            
            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def _retrieve_status_report_view(self, arguments: CliArgument)->None:
        
        try:
            data_frame = self.get_status_report_data_frame(arguments)

            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
       
    def _retrieve_asset_inventory_view(self, arguments: CliArgument)->None:
        
        try:
            data_frame = self.get_asset_inventory_data_frame(arguments)

            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
    
    def _retrieve_project_view(self, arguments: CliArgument)->None:
        
        try:
            data_frame = self.get_project_data_frame(arguments)

            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
       
    def _send_to_output(self, data_frame: pd.DataFrame, format, out_file, view) -> None:
        
        if out_file is None:
            file = DEFAULT_FILENAME
            out_file = file
        
        out_file = append_date_to_file_name(out_file)
        try:
            if format == OUTPUT_FORMAT[1]:
                json_file = f"{file}.json" if out_file is None else out_file
                data_frame.to_json(json_file, orient='records')
                print(f"the file was saved : {json_file}")
            elif format == OUTPUT_FORMAT[2]:
                csv_file = f"{file}.csv" if out_file is None else out_file
                data_frame.to_csv(csv_file)
                print(f"the file was saved : {csv_file}")
            elif format == OUTPUT_FORMAT[3]:
                excel_file = f"{file}.xlsx" if out_file is None else out_file

                if view == SUPPORTED_VIEWS[0]: # status-report
                    self._generate_raw_status_report_excel(
                        data_frame, 
                        excel_file
                    )
                elif view == SUPPORTED_VIEWS[1]: #asset-inventory
                    self._generate_raw_asset_inventory_excel(
                        data_frame, 
                        excel_file
                    )
                elif view == SUPPORTED_VIEWS[2]: #project
                    self._generate_raw_excel(
                        data_frame, 
                        excel_file,
                        'Projects'
                    )
                else:
                    self._log(f'error = {NOT_SUPPORTED}', 1)
                
            elif format == OUTPUT_FORMAT[0]:    
                print('\n') 
                print(tabulate(
                    data_frame, 
                    showindex=False, 
                    headers=data_frame.columns
                ))
                print('\n') 
            else:
                print(NOT_SUPPORTED)

        except Exception as ex:
            self._log(str(ex),2)

    def _generate_raw_status_report_excel(self, data_frame : pd.DataFrame, file_name)-> None:
        try:
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)

            drop_columns=['WeekRangeStart', 'WeekRangeId', 'ProjectId']
            column_headers = [
                'Project',
                'Project Code',
                'Rework',
                'Ref ID',
                'Description',
                'Severity',
                'Incident Type',
                'Assigned Employee',
                'Phase',
                'Status',
                'Start Date',
                'Target Date',
                'Completion Date',
                'Estimate',
                'Week Effort',
                'Actual Effort',
                'Comments',
                'Inbound Contacts',
                'Week End Date',
                'DepartmentID',
                'DivsionID'

            ]

            data_frame.drop(drop_columns, inplace=True, axis=1)
            data_frame.columns = column_headers
            total_row = len(data_frame)

            grouped_per_project = data_frame.groupby('Project')

            for key, values in grouped_per_project:
                df_per_group = pd.DataFrame(grouped_per_project.get_group(key))
                #df_per_group.columns = column_headers
                grouped_per_project_by_week = df_per_group.groupby([
                    'Week End Date'
                ])
               
                df_per_group.to_excel(
                    self._writer, 
                    sheet_name=key, 
                    index= False 
                )

                worksheet = self._writer.sheets[key]
                worksheet.set_column(0,3,12)
                worksheet.set_column(4,4,45)
                worksheet.set_column(5,5,12.5)
                worksheet.set_column(6,6,20)
                worksheet.set_column(7,7,20)
                worksheet.set_column(8,12,18)
                worksheet.set_column(13,15,12)
                worksheet.set_column(16,19,25)
                
                for col_num, value in enumerate(df_per_group.columns.values):
                    worksheet.write(0, col_num, value, main_header_format)       

            worksheet.write(
                total_row + 4, 
                0, 
                f'Report generated : {datetime.now()}', 
                footer_format
            )
                 
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _generate_raw_asset_inventory_excel(self, data_frame : pd.DataFrame, file_name)-> None:
        try:
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)

            drop_columns=[
                'Comments', 
                'OtherInfo', 
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
                'Asset Tag'
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
            print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _generate_raw_asset_inventory_excel(self, data_frame : pd.DataFrame, file_name)-> None:
        try:
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)

        
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
            print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _generate_raw_excel(self, data_frame : pd.DataFrame, file_name, sheet_name = 'Sheet1')-> None:
        try:
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
        
            data_frame.to_excel(
                self._writer, 
                sheet_name=sheet_name, 
                index= False 
            )

            worksheet = self._writer.sheets[sheet_name]
            
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
            print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)
 
    def _load_report_settings(self) -> None:
        self._report_settings = self._config.get_section('reports')
        self._report_schedules = self._report_settings['schedules']

    def _get_report_schedule(self, month: int) -> List[str]:
        return self._report_schedules[month-1]
