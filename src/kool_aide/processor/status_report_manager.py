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


from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *

from kool_aide.db_access.connection import Connection
from kool_aide.db_access.dbhelper.status_report_helper \
    import StatusReportHelper

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
        
        self._db_helper = StatusReportHelper(
            self._logger, 
            self._config, 
            self._connection
        )

        self._report_settings = None
        self._report_schedules = None

        self._load_report_settings()

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
            
            # if running as auto, use settings. else, use params
            if arguments.auto_mode:
                week_filter = self._get_report_schedule(datetime.now().month)

            results = self._db_helper.get_status_report_view(week_filter)
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if project_filter is not None and len(project_filter)>0:
                data_frame = data_frame[data_frame['Project'].isin(project_filter)]
            
            
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
    
    def _retrieve_status_report_view(self, arguments: CliArgument)->None:
        
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
       
    def _send_to_output(self, data_frame: pd.DataFrame, format, out_file) ->None:
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

    def _process_excel(self, data_frame : pd.DataFrame, file_name)-> None:
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
            
            workbook = writer.book
            main_header_format = workbook.add_format(SHEET_TOP_HEADER)
            header_format_orange = workbook.add_format(SHEET_HEADER_ORANGE)
            header_format_gray = workbook.add_format(SHEET_HEADER_GRAY)
            cell_wrap_noborder = workbook.add_format(SHEET_CELL_WRAP_NOBORDER)
            cell_wrap_noborder_alt = workbook.add_format(SHEET_CELL_WRAP_NOBORDER_ALT)
            cell_total = workbook.add_format(SHEET_HEADER_LT_GREEN)
            cell_sub_total = workbook.add_format(SHEET_HEADER_LT_BLUE)

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

                worksheet = writer.sheets[key]
                worksheet.set_column(0,3,12,cell_wrap_noborder)
                worksheet.set_column(4,4,45, cell_wrap_noborder)
                worksheet.set_column(5,5,12.5)
                worksheet.set_column(6,6,20)
                worksheet.set_column(7,7,20)
                worksheet.set_column(8,12,18)
                worksheet.set_column(13,15,12)
                worksheet.set_column(16,19,25)
                

                for col_num, value in enumerate(df_per_group.columns.values):
                    worksheet.write(0, col_num, value, main_header_format)       

                total_row = len(df_per_group)
                total_hrs = sum(df_per_group['Week Effort'])
                
                # worksheet.write(total_row + 3, 0, f'Total Entries')
                # worksheet.write(total_row + 3, 1, f'{total_row}')
                
                grouped_by_week = df_per_group.groupby([
                    'Week End Date'
                ])
                # print(grouped_by_phase_status['Week Effort'].sum())
                worksheet.write(total_row + 3, 0, f'Weekly Time Entries', header_format_orange)
                worksheet.write(total_row + 3, 1, f'', header_format_orange)
                worksheet.write(total_row + 4, 0, f'Week Ending', header_format_gray)
                worksheet.write(total_row + 4, 1, f'Hours', header_format_gray)
                index = 5
                for group, value in grouped_by_week['Week Effort'].sum().items():
                    worksheet.write(total_row + index, 0, f'{group}')
                    worksheet.write_number(total_row + index, 1, value)
                    index += 1
                    #print(f'{group}  - {value}')
                    # pass
                worksheet.write(total_row + index, 0, f'Total Hours', header_format_orange)
                worksheet.write_number(total_row + index, 1, total_hrs, cell_total)

                grouped_by_week_per_employee_by_incident_type = \
                    df_per_group.groupby([
                        'Week End Date',
                        'Assigned Employee',
                        'Incident Type'
                    ])

                worksheet.write(total_row + 3, 3, f'Time Entries By Employees Per Incident Type Per Week', header_format_orange)
                worksheet.write(total_row + 3, 4, f'', header_format_orange)
                worksheet.write(total_row + 3, 5, f'', header_format_orange)
                worksheet.write(total_row + 3, 6, f'', header_format_orange)
                for group, value in grouped_by_week_per_employee_by_incident_type.sum().items():
                    
                    if group=='Week Effort':
                        worksheet.write(total_row + 4, 3, f'Week Ending', header_format_gray)
                        worksheet.write(total_row + 4, 4, f'Assigned Employee', header_format_gray)
                        worksheet.write(total_row + 4, 5, f'Incident Type', header_format_gray)
                        worksheet.write(total_row + 4, 6, f'Hours', header_format_gray)
                        index = 5
                        date_list=[]
                        emp_list=[]
                        add_date= False
                        cell_format = cell_wrap_noborder
                        group_index = 1
                        weekly_hours = 0
                        for inner_group, inner_value in value.items():
                            #print(f'{g1[2]} : {v1}')
                            if len(date_list) == 0:
                                date_list.append(inner_group[0])
                                add_date = True
                            else:
                                if inner_group[0] in date_list:
                                    add_date = False
                                else:
                                    date_list.append(inner_group[0])
                                    add_date = True
                                    emp_list.clear()
                                    group_index += 1
                                    worksheet.write_number(total_row + index, 6, weekly_hours, cell_sub_total)
                                    weekly_hours = 0
                                    index += 1


                            if add_date:
                                worksheet.write(total_row + index, 3,f'{inner_group[0]}', cell_format)
                                worksheet.write(total_row + index, 4,f'{inner_group[1]}', cell_format)
                                emp_list.append(inner_group[1])
                            else:
                                if inner_group[1] not in emp_list:
                                    worksheet.write(total_row + index, 4,f'{inner_group[1]}', cell_format)
                                    emp_list.append(inner_group[1])

                            worksheet.write(total_row + index, 5,f'{inner_group[2]}', cell_format)
                            worksheet.write_number(total_row + index, 6,inner_value, cell_format)
                            weekly_hours += inner_value
                            index += 1

                        worksheet.write_number(total_row + index, 6, weekly_hours, cell_sub_total)
                        weekly_hours = 0
                        index += 1
                        worksheet.write(total_row + index, 3, f'Total Hours', header_format_orange)
                        worksheet.write(total_row + index, 4, '', header_format_orange)
                        worksheet.write(total_row + index, 5, '', header_format_orange)
                        worksheet.write_number(total_row + index, 6, total_hrs, cell_total)
                    
            writer.save()

            # data_frame.to_excel(file_name, index=False)
            print(f"the file was saved : {file_name}") 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _load_report_settings(self) -> None:
        self._report_settings = self._config.get_section('reports')
        self._report_schedules = self._report_settings['schedules']

    def _get_report_schedule(self, month: int) -> List[str]:
        return self._report_schedules[month-1]
