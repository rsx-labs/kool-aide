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


from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *

from kool_aide.db_access.connection import Connection
from kool_aide.db_access.dbhelper.status_report_helper import StatusReportHelper

from kool_aide.model.cli_argument import CliArgument
from kool_aide.model.aide.project import Project
from kool_aide.model.aide.week_range import WeekRange
from kool_aide.model.aide.status_report import StatusReport

from kool_aide.assets.resources.messages import *


class StatusReportManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument = None):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments
        self._db_helper = StatusReportHelper(self._logger, self._config, self._connection)
        self._report_settings = self._load_report_settings()

        self._log("initialize")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [processor.status_report_manager]", level)

    def retrieve(self, arguments : CliArgument):
        self._log(f"retrieving model : {arguments.model}")
  
        if arguments.model == SUPPORTED_MODELS[4]:
            self._retrieve_status_report_view(arguments)
            return True, DATA_RETRIEVED
        
        return False, NOT_SUPPORTED

    def get_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        project_filter = None
        week_filter = None
        
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    project_filter = None if PARAM_PROJECT not in json_parameters else json_parameters[PARAM_PROJECT] 
                    week_filter = None if PARAM_WEEK not in json_parameters else json_parameters[PARAM_WEEK] 
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_status_report_view(week_filter)
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()


            if project_filter is not None and len(project_filter)>0:
                data_frame = data_frame[data_frame['Project'].isin(project_filter)]
            
            # if week_filter is not None and len(week_filter)>0:
            #     data_frame = data_frame[data_frame['WeekRangeId'].isin(week_filter)]
            
            if sort_keys is not None and len(sort_keys)>0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns)>0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
            
            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def _retrieve_status_report_view(self, arguments: CliArgument):
        
        try:
            data_frame = self.get_data_frame(arguments)

            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
       
    def _send_to_output(self, data_frame: pd.DataFrame, format, out_file):
        if out_file is None:
            file = DEFAULT_FILENAME
        try:
            if format == DISPLAY_FORMAT[1]:
                json_file = f"{file}.json" if out_file is None else out_file
                data_frame.to_json(json_file, orient='records')
                print(f"the file was saved : {json_file}")
            elif format == DISPLAY_FORMAT[2]:
                csv_file = f"{file}.csv" if out_file is None else out_file
                data_frame.to_csv(csv_file)
                print(f"the file was saved : {csv_file}")
            elif format == DISPLAY_FORMAT[3]:
                excel_file = f"{file}.xslx" if out_file is None else out_file
                self._process_excel(data_frame, excel_file)
                
            elif format == DISPLAY_FORMAT[0]:    
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

    def _process_excel(self, data_frame : pd.DataFrame, file_name):
        try:
            # check if the filename contains placeholder [D]

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
            ]

            data_frame.drop(drop_columns, inplace=True, axis=1)
            if '[D]' in file_name:
                date_now = datetime.now()
                date_string = f'{MONTHS[date_now.month-1]}{date_now.year}'
                file_name = file_name.replace('[D]',date_string)

            grouped_per_project = data_frame.groupby('Project')

            # Create a Pandas Excel writer using XlsxWriter as the engine.
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            for key, values in grouped_per_project:
                df_per_group = pd.DataFrame(grouped_per_project.get_group(key))
                df_per_group.columns = column_headers
                grouped_per_project_by_week = df_per_group.groupby([
                    'Week End Date'
                ])
               
                df_per_group.to_excel(
                    writer, 
                    sheet_name=key, 
                    index= False 
                )

                workbook = writer.book
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 0
                })
                sub1_format = workbook.add_format({
                    'bold': True,
                    'text_wrap':False,
                    'valign': 'top',
                    'fg_color': '#FDBF42',
                    'border': 0
                })
                sub2_format = workbook.add_format({
                    'bold': True,
                    'text_wrap':False,
                    'valign': 'top',
                    'fg_color': '#BBB9B5',
                    'border': 0
                })


                worksheet = writer.sheets[key]
                worksheet.set_column(0,3,12)
                worksheet.set_column(4,4,45)
                worksheet.set_column(6,6,20)
                worksheet.set_column(7,7,20)
                worksheet.set_column(8,12,18)
                worksheet.set_column(13,15,12)
                worksheet.set_column(16,19,25)

                for col_num, value in enumerate(df_per_group.columns.values):
                    worksheet.write(0, col_num, value, header_format)       

                total_row = len(df_per_group)
                total_hrs = sum(df_per_group['Week Effort'])
                
                # worksheet.write(total_row + 3, 0, f'Total Entries')
                # worksheet.write(total_row + 3, 1, f'{total_row}')
                
                grouped_by_phase_status = df_per_group.groupby([
                    'Week End Date'
                ])
                #print(grouped_by_phase_status['Week Effort'].sum())
                worksheet.write(total_row + 3, 0, f'Time Entries By Week', sub1_format)
                worksheet.write(total_row + 3, 1, f'', sub1_format)
                worksheet.write(total_row + 3, 2, f'', sub1_format)
                worksheet.write(total_row + 4, 0, f'Week Ending', sub2_format)
                worksheet.write(total_row + 5, 0, f'Hours', sub2_format)
                index = 1
                for group, value in grouped_by_phase_status['Week Effort'].sum().items():
                    worksheet.write(total_row + 4, index, f'{group}')
                    worksheet.write(total_row + 5, index, f'{value}')
                    index += 1
                    #print(f'{group}  - {value}')
                    # pass
                worksheet.write(total_row + 6 + 1, 0, f'Total Hours', sub1_format)
                worksheet.write(total_row + 6 + 1, 1, f'{total_hrs}')

            writer.save()

            # data_frame.to_excel(file_name, index=False)
            print(f"the file was saved : {file_name}") 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _load_report_settings(self):
        pass