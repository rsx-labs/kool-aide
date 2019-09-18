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

        self._writer = None
        self._workbook = None
        self._main_header_format = None
        self._header_format_orange = None
        self._header_format_gray = None
        self._cell_wrap_noborder = None
        self._cell_total = None
        self._cell_sub_total = None
        
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
       
    def _send_to_output(self, data_frame: pd.DataFrame, format, out_file) -> None:
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
                excel_file = f"{file}.xlsx" if out_file is None else out_file

                if '[D]' in excel_file:
                    date_now = datetime.now()
                    date_string = f'{MONTHS[date_now.month-1]}{date_now.year}'
                    excel_file = excel_file.replace('[D]',date_string)

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
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            self._main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            self._header_format_orange = self._workbook.add_format(SHEET_HEADER_ORANGE)
            self._header_format_gray = self._workbook.add_format(SHEET_HEADER_GRAY)
            self._cell_wrap_noborder = self._workbook.add_format(SHEET_CELL_WRAP_NOBORDER)
            self._cell_total = self._workbook.add_format(SHEET_HEADER_LT_GREEN)
            self._cell_sub_total = self._workbook.add_format(SHEET_HEADER_GAINSBORO)

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
            data_frame.columns = column_headers

            self._create_main_report(data_frame)
            self._create_summary_by_week_report(data_frame)
            self._create_summary_report(data_frame)
                          
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            print(f"the file was saved : {file_name}") 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_main_report(self, data_frame : pd.DataFrame)-> None:
        try:
            
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
                worksheet.set_column(0,3,12,self._cell_wrap_noborder)
                worksheet.set_column(4,4,45, self._cell_wrap_noborder)
                worksheet.set_column(5,5,12.5, self._cell_wrap_noborder)
                worksheet.set_column(6,6,20, self._cell_wrap_noborder)
                worksheet.set_column(7,7,20, self._cell_wrap_noborder)
                worksheet.set_column(8,12,18, self._cell_wrap_noborder)
                worksheet.set_column(13,15,12, self._cell_wrap_noborder)
                worksheet.set_column(16,19,25, self._cell_wrap_noborder)
                
                for col_num, value in enumerate(df_per_group.columns.values):
                    worksheet.write(0, col_num, value, self._main_header_format)       

                total_row = len(df_per_group)
                total_hrs = sum(df_per_group['Week Effort'])
         
                grouped_by_week = df_per_group.groupby([
                    'Week End Date'
                ])
                # print(grouped_by_phase_status['Week Effort'].sum())
                worksheet.write(total_row + 3, 0, f'Weekly Time Entries', self._header_format_orange)
                worksheet.write(total_row + 3, 1, f'', self._header_format_orange)
                worksheet.write(total_row + 4, 0, f'Week Ending', self._header_format_gray)
                worksheet.write(total_row + 4, 1, f'Hours', self._header_format_gray)
                index = 5
                for group, value in grouped_by_week['Week Effort'].sum().items():
                    worksheet.write(total_row + index, 0, f'{group}')
                    worksheet.write_number(total_row + index, 1, value)
                    index += 1
                   
                worksheet.write(total_row + index, 0, f'Total Hours', self._header_format_orange)
                worksheet.write_number(total_row + index, 1, total_hrs, self._cell_total)

                grouped_by_week_per_employee_by_incident_type = \
                    df_per_group.groupby([
                        'Week End Date',
                        'Assigned Employee',
                        'Incident Type'
                    ])

                worksheet.write(total_row + 3, 3, 'Time Entries By Employees Per Incident Type Per Week', self._header_format_orange)
                worksheet.write(total_row + 3, 4, '', self._header_format_orange)
                worksheet.write(total_row + 3, 5, '', self._header_format_orange)
                worksheet.write(total_row + 3, 6, '', self._header_format_orange)
                for group, value in grouped_by_week_per_employee_by_incident_type.sum().items():
                    
                    if group=='Week Effort':
                        worksheet.write(total_row + 4, 3, 'Week Ending', self._header_format_gray)
                        worksheet.write(total_row + 4, 4, 'Assigned Employee', self._header_format_gray)
                        worksheet.write(total_row + 4, 5, 'Incident Type', self._header_format_gray)
                        worksheet.write(total_row + 4, 6, 'Hours', self._header_format_gray)
                        index = 5
                        date_list=[]
                        emp_list=[]
                        add_date= False
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
                                    worksheet.write(total_row + index, 3, '', self._cell_sub_total)
                                    worksheet.write(total_row + index, 4, '', self._cell_sub_total)
                                    worksheet.write(total_row + index, 5, '', self._cell_sub_total)
                                    worksheet.write_number(total_row + index, 6, weekly_hours, self._cell_sub_total)
                                    weekly_hours = 0
                                    index += 1


                            if add_date:
                                worksheet.write(total_row + index, 3,f'{inner_group[0]}', self._cell_wrap_noborder)
                                worksheet.write(total_row + index, 4,f'{inner_group[1]}', self._cell_wrap_noborder)
                                emp_list.append(inner_group[1])
                            else:
                                if inner_group[1] not in emp_list:
                                    worksheet.write(total_row + index, 4,f'{inner_group[1]}', self._cell_wrap_noborder)
                                    emp_list.append(inner_group[1])

                            worksheet.write(total_row + index, 5,f'{inner_group[2]}', self._cell_wrap_noborder)
                            worksheet.write_number(total_row + index, 6,inner_value, self._cell_wrap_noborder)
                            weekly_hours += inner_value
                            index += 1

                        worksheet.write(total_row + index, 3, '', self._cell_sub_total)
                        worksheet.write(total_row + index, 4, '', self._cell_sub_total)
                        worksheet.write(total_row + index, 5, '', self._cell_sub_total)
                        worksheet.write_number(total_row + index, 6, weekly_hours, self._cell_sub_total)
                        weekly_hours = 0
                        index += 1
                        worksheet.write(total_row + index, 3, f'Total Hours', self._header_format_orange)
                        worksheet.write(total_row + index, 4, '', self._header_format_orange)
                        worksheet.write(total_row + index, 5, '', self._header_format_orange)
                        worksheet.write_number(total_row + index, 6, total_hrs, self._cell_total)
       
        except Exception as ex:
            self._log(f'error generating main report. {str(ex)}', 2)

    def _create_summary_by_week_report(self, data_frame : pd.DataFrame) -> None:

        worksheet = self._workbook.add_worksheet('Team Summary By Week')
        grouped_per_project = data_frame.groupby(['Week End Date', 'Project'])

        total_row = 0
        worksheet.write(total_row, 0, 'Time Entries Per Project By Week', self._header_format_orange)
        worksheet.write(total_row, 1, '', self._header_format_orange)
        worksheet.write(total_row, 2, '', self._header_format_orange)
        worksheet.write(total_row + 1, 0, 'Week Ending', self._header_format_gray)
        worksheet.write(total_row + 1, 1, 'Project', self._header_format_gray)
        worksheet.write(total_row + 1, 2, 'Hours', self._header_format_gray)

        worksheet.set_column(0,0,12,self._cell_wrap_noborder)
        worksheet.set_column(1,1,18,self._cell_wrap_noborder)
        worksheet.set_column(2,2,9,self._cell_wrap_noborder)
        worksheet.set_column(3,3,5,self._cell_wrap_noborder)
        date_list=[]
        sub_list = []
        add_date= False
        weekly_hours = 0
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        for key, value in grouped_per_project['Week Effort'].sum().items():
           
            #print(f'{g1[2]} : {v1}')
            if len(date_list) == 0:
                date_list.append(key[0])
                add_date = True
            else:
                if key[0] in date_list:
                    add_date = False
                else:
                    date_list.append(key[0])
                    sub_list.clear()
                    add_date = True
                    worksheet.write(total_row + index, 0, '', self._cell_sub_total)
                    worksheet.write(total_row + index, 1, '', self._cell_sub_total)
                    worksheet.write_number(total_row + index, 2, weekly_hours, self._cell_sub_total)
                    weekly_hours = 0
                    index += 1

            if add_date:
                worksheet.write(total_row + index, 0,f'{key[0]}', self._cell_wrap_noborder)
                worksheet.write(total_row + index, 1,f'{key[1]}', self._cell_wrap_noborder)
            else:
                worksheet.write(total_row + index, 1,f'{key[1]}', self._cell_wrap_noborder)
            
            worksheet.write_number(total_row + index, 2,value, self._cell_wrap_noborder)
            weekly_hours += value
            index += 1
        worksheet.write(total_row + index, 0, '', self._cell_sub_total)
        worksheet.write(total_row + index, 1, '', self._cell_sub_total)
        worksheet.write_number(total_row + index, 2, weekly_hours, self._cell_sub_total)
        weekly_hours = 0
        index += 1
        worksheet.write(total_row + index, 0, f'Total Hours', self._header_format_orange)
        worksheet.write(total_row + index, 1, '', self._header_format_orange)
        worksheet.write_number(total_row + index, 2, total_hrs, self._cell_total)

        grouped_per_project_by_incident = data_frame.groupby(['Week End Date', 'Project','Incident Type'])

        worksheet.write(total_row, 4, 'Time Entries Per Project By Incident By Week', self._header_format_orange)
        worksheet.write(total_row, 5, '', self._header_format_orange)
        worksheet.write(total_row, 6, '', self._header_format_orange)
        worksheet.write(total_row, 7, '', self._header_format_orange)
        worksheet.write(total_row + 1, 4, 'Week Ending', self._header_format_gray)
        worksheet.write(total_row + 1, 5, 'Project', self._header_format_gray)
        worksheet.write(total_row + 1, 6, 'Incident Type', self._header_format_gray)
        worksheet.write(total_row + 1, 7, 'Hours', self._header_format_gray)

        worksheet.set_column(4,4,12,self._cell_wrap_noborder)
        worksheet.set_column(5,5,25,self._cell_wrap_noborder)
        worksheet.set_column(6,6,15,self._cell_wrap_noborder)
        worksheet.set_column(7,7,9,self._cell_wrap_noborder)
        worksheet.set_column(8,8,5,self._cell_wrap_noborder)

        date_list=[]
        sub_list=[]
        add_date= False
        weekly_hours = 0
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        for key, value in grouped_per_project_by_incident['Week Effort'].sum().items():
           
            #print(f'{g1[2]} : {v1}')
            if len(date_list) == 0:
                date_list.append(key[0])
                add_date = True
            else:
                if key[0] in date_list:
                    add_date = False
                else:
                    date_list.append(key[0])
                    sub_list.clear()
                    add_date = True
                    worksheet.write(total_row + index, 4, '', self._cell_sub_total)
                    worksheet.write(total_row + index, 5, '', self._cell_sub_total)
                    worksheet.write(total_row + index, 6, '', self._cell_sub_total)
                    worksheet.write_number(total_row + index, 7, weekly_hours, self._cell_sub_total)
                    weekly_hours = 0
                    index += 1

            if add_date:
                worksheet.write(total_row + index, 4,f'{key[0]}', self._cell_wrap_noborder)
                worksheet.write(total_row + index, 5,f'{key[1]}', self._cell_wrap_noborder)
                sub_list.append(key[1])    
            else:
                if key[1] not in sub_list:
                    worksheet.write(total_row + index, 5,f'{key[1]}', self._cell_wrap_noborder)
                    sub_list.append(key[1])
            
            worksheet.write(total_row + index, 6,f'{key[2]}', self._cell_wrap_noborder)
            worksheet.write_number(total_row + index, 7,value, self._cell_wrap_noborder)
            weekly_hours += value
            index += 1

        worksheet.write(total_row + index, 4, '', self._cell_sub_total)
        worksheet.write(total_row + index, 5, '', self._cell_sub_total)
        worksheet.write(total_row + index, 6, '', self._cell_sub_total)
        worksheet.write_number(total_row + index, 7, weekly_hours, self._cell_sub_total)
        weekly_hours = 0
        index += 1
        worksheet.write(total_row + index, 4, f'Total Hours', self._header_format_orange)
        worksheet.write(total_row + index, 5, '', self._header_format_orange)
        worksheet.write(total_row + index, 6, '', self._header_format_orange)
        worksheet.write_number(total_row + index, 7, total_hrs, self._cell_total)

        #######
        grouped_per_project = data_frame.groupby(['Week End Date', 'Assigned Employee'])

        total_row = 0
        worksheet.write(total_row, 9, 'Time Entries Per Employee By Week', self._header_format_orange)
        worksheet.write(total_row, 10, '', self._header_format_orange)
        worksheet.write(total_row, 11, '', self._header_format_orange)
        worksheet.write(total_row + 1, 9, 'Week Ending', self._header_format_gray)
        worksheet.write(total_row + 1, 10, 'Employee', self._header_format_gray)
        worksheet.write(total_row + 1, 11, 'Hours', self._header_format_gray)

        worksheet.set_column(9,9,12,self._cell_wrap_noborder)
        worksheet.set_column(10,10,25,self._cell_wrap_noborder)
        worksheet.set_column(11,11,9,self._cell_wrap_noborder)
        worksheet.set_column(12,12,5,self._cell_wrap_noborder)
        date_list=[]
        sub_list = []
        add_date= False
        weekly_hours = 0
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        for key, value in grouped_per_project['Week Effort'].sum().items():
           
            #print(f'{g1[2]} : {v1}')
            if len(date_list) == 0:
                date_list.append(key[0])
                add_date = True
            else:
                if key[0] in date_list:
                    add_date = False
                else:
                    date_list.append(key[0])
                    sub_list.clear()
                    add_date = True
                    worksheet.write(total_row + index, 9, '', self._cell_sub_total)
                    worksheet.write(total_row + index, 10, '', self._cell_sub_total)
                    worksheet.write_number(total_row + index, 11, weekly_hours, self._cell_sub_total)
                    weekly_hours = 0
                    index += 1

            if add_date:
                worksheet.write(total_row + index, 9,f'{key[0]}', self._cell_wrap_noborder)
                worksheet.write(total_row + index, 10,f'{key[1]}', self._cell_wrap_noborder)
            else:
                worksheet.write(total_row + index, 10,f'{key[1]}', self._cell_wrap_noborder)
            
            worksheet.write_number(total_row + index, 11,value, self._cell_wrap_noborder)
            weekly_hours += value
            index += 1
        worksheet.write(total_row + index, 9, '', self._cell_sub_total)
        worksheet.write(total_row + index, 10, '', self._cell_sub_total)
        worksheet.write_number(total_row + index, 11, weekly_hours, self._cell_sub_total)
        weekly_hours = 0
        index += 1
        worksheet.write(total_row + index, 9, f'Total Hours', self._header_format_orange)
        worksheet.write(total_row + index, 10, '', self._header_format_orange)
        worksheet.write_number(total_row + index, 11, total_hrs, self._cell_total)

        grouped_per_project_by_incident = data_frame.groupby(['Week End Date', 'Assigned Employee','Project'])

        worksheet.write(total_row, 13, 'Time Entries Per Employee By Project By Week', self._header_format_orange)
        worksheet.write(total_row, 14, '', self._header_format_orange)
        worksheet.write(total_row, 15, '', self._header_format_orange)
        worksheet.write(total_row, 16, '', self._header_format_orange)
        worksheet.write(total_row + 1, 13, 'Week Ending', self._header_format_gray)
        worksheet.write(total_row + 1, 14, 'Assigned Employee', self._header_format_gray)
        worksheet.write(total_row + 1, 15, 'Project', self._header_format_gray)
        worksheet.write(total_row + 1, 16, 'Hours', self._header_format_gray)

        worksheet.set_column(13,13,12,self._cell_wrap_noborder)
        worksheet.set_column(14,14,25,self._cell_wrap_noborder)
        worksheet.set_column(15,15,25,self._cell_wrap_noborder)
        worksheet.set_column(16,16,9,self._cell_wrap_noborder)
        worksheet.set_column(17,17,5,self._cell_wrap_noborder)

        date_list=[]
        sub_list=[]
        add_date= False
        weekly_hours = 0
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        for key, value in grouped_per_project_by_incident['Week Effort'].sum().items():
           
            #print(f'{g1[2]} : {v1}')
            if len(date_list) == 0:
                date_list.append(key[0])
                add_date = True
            else:
                if key[0] in date_list:
                    add_date = False
                else:
                    date_list.append(key[0])
                    sub_list.clear()
                    add_date = True
                    worksheet.write(total_row + index, 13, '', self._cell_sub_total)
                    worksheet.write(total_row + index, 14, '', self._cell_sub_total)
                    worksheet.write(total_row + index, 15, '', self._cell_sub_total)
                    worksheet.write_number(total_row + index, 16, weekly_hours, self._cell_sub_total)
                    weekly_hours = 0
                    index += 1

            if add_date:
                worksheet.write(total_row + index, 13,f'{key[0]}', self._cell_wrap_noborder)
                worksheet.write(total_row + index, 14,f'{key[1]}', self._cell_wrap_noborder)
                sub_list.append(key[1])    
            else:
                if key[1] not in sub_list:
                    worksheet.write(total_row + index, 14,f'{key[1]}', self._cell_wrap_noborder)
                    sub_list.append(key[1])
            
            worksheet.write(total_row + index, 15,f'{key[2]}', self._cell_wrap_noborder)
            worksheet.write_number(total_row + index, 16,value, self._cell_wrap_noborder)
            weekly_hours += value
            index += 1

        worksheet.write(total_row + index, 13, '', self._cell_sub_total)
        worksheet.write(total_row + index, 14, '', self._cell_sub_total)
        worksheet.write(total_row + index, 15, '', self._cell_sub_total)
        worksheet.write_number(total_row + index, 16, weekly_hours, self._cell_sub_total)
        weekly_hours = 0
        index += 1
        worksheet.write(total_row + index, 13, f'Total Hours', self._header_format_orange)
        worksheet.write(total_row + index, 14, '', self._header_format_orange)
        worksheet.write(total_row + index, 15, '', self._header_format_orange)
        worksheet.write_number(total_row + index, 16, total_hrs, self._cell_total)

    def _create_summary_report(self, data_frame : pd.DataFrame) -> None:

        worksheet = self._workbook.add_worksheet('Team Summary')
        grouped_per_project = data_frame.groupby(['Project'])

        total_row = 0
        worksheet.write(total_row, 0, 'Time Entries Per Project', self._header_format_orange)
        worksheet.write(total_row, 1, '', self._header_format_orange)
        worksheet.write(total_row, 2, '', self._header_format_orange)
        worksheet.write(total_row + 1, 0, 'Project', self._header_format_gray)
        worksheet.write(total_row + 1, 1, 'Hours', self._header_format_gray)

        worksheet.set_column(0,0,20,self._cell_wrap_noborder)
        worksheet.set_column(1,1,9,self._cell_wrap_noborder)
        worksheet.set_column(2,2,5,self._cell_wrap_noborder)
        project_list=[]
        sub_list = []
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        for key, value in grouped_per_project['Week Effort'].sum().items():
            worksheet.write(total_row + index, 0,f'{key}', self._cell_wrap_noborder)
            worksheet.write_number(total_row + index, 1,value, self._cell_wrap_noborder)
            index += 1
        
        worksheet.write(total_row + index, 0, f'Total Hours', self._header_format_orange)
        worksheet.write_number(total_row + index, 1, total_hrs, self._cell_total)

        ######

        grouped_per_project_by_incident = data_frame.groupby(['Project','Incident Type'])

        worksheet.write(total_row, 3, 'Time Entries Per Project By Incident', self._header_format_orange)
        worksheet.write(total_row, 4, '', self._header_format_orange)
        worksheet.write(total_row, 5, '', self._header_format_orange)
        worksheet.write(total_row + 1, 3, 'Project', self._header_format_gray)
        worksheet.write(total_row + 1, 4, 'Incident Type', self._header_format_gray)
        worksheet.write(total_row + 1, 5, 'Hours', self._header_format_gray)

        worksheet.set_column(3,3,25,self._cell_wrap_noborder)
        worksheet.set_column(4,4,15,self._cell_wrap_noborder)
        worksheet.set_column(5,5,9,self._cell_wrap_noborder)
        worksheet.set_column(6,6,5,self._cell_wrap_noborder)

        project_list=[]
        sub_list=[]
        add_project= False
        weekly_hours = 0
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        for key, value in grouped_per_project_by_incident['Week Effort'].sum().items():
           
            #print(f'{g1[2]} : {v1}')
            if len(project_list) == 0:
                project_list.append(key[0])
                add_project= True
            else:
                if key[0] in project_list:
                    add_project = False
                else:
                    project_list.append(key[0])
                    sub_list.clear()
                    add_project = True
                    worksheet.write(total_row + index, 3, '', self._cell_sub_total)
                    worksheet.write(total_row + index, 4, '', self._cell_sub_total)
                    worksheet.write_number(total_row + index, 5, weekly_hours, self._cell_sub_total)
                    weekly_hours = 0
                    index += 1

            if add_project:
                worksheet.write(total_row + index, 3,f'{key[0]}', self._cell_wrap_noborder)
                worksheet.write(total_row + index, 4,f'{key[1]}', self._cell_wrap_noborder)
            else:
                worksheet.write(total_row + index, 4,f'{key[1]}', self._cell_wrap_noborder)
                   
            worksheet.write_number(total_row + index, 5,value, self._cell_wrap_noborder)
            weekly_hours += value
            index += 1

        worksheet.write(total_row + index, 3, '', self._cell_sub_total)
        worksheet.write(total_row + index, 4, '', self._cell_sub_total)
        worksheet.write_number(total_row + index, 5, weekly_hours, self._cell_sub_total)
        weekly_hours = 0
        index += 1
        worksheet.write(total_row + index, 3, f'Total Hours', self._header_format_orange)
        worksheet.write(total_row + index, 4, '', self._header_format_orange)
        worksheet.write_number(total_row + index, 5, total_hrs, self._cell_total)

        #######
        grouped_per_employee= data_frame.groupby(['Assigned Employee'])

        total_row = 0
        worksheet.write(total_row, 7, 'Time Entries Per Employee', self._header_format_orange)
        worksheet.write(total_row, 8, '', self._header_format_orange)
        worksheet.write(total_row + 1, 7, 'Employee', self._header_format_gray)
        worksheet.write(total_row + 1, 8, 'Hours', self._header_format_gray)
       
        worksheet.set_column(7,7,25,self._cell_wrap_noborder)
        worksheet.set_column(8,8,9,self._cell_wrap_noborder)
        worksheet.set_column(9,9,5,self._cell_wrap_noborder)
        project_list=[]
        sub_list = []
        add_project= False
        weekly_hours = 0
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        for key, value in grouped_per_employee['Week Effort'].sum().items():
            worksheet.write(total_row + index, 7, f'{key}', self._cell_wrap_noborder)
            worksheet.write_number(total_row + index, 8, value, self._cell_wrap_noborder)
            index +=1
        worksheet.write(total_row + index, 7, f'Total Hours', self._header_format_orange)
        worksheet.write_number(total_row + index, 8, total_hrs, self._cell_total)

        ######

        grouped_per_employee_by_project = data_frame.groupby(['Assigned Employee','Project'])

        worksheet.write(total_row, 10, 'Time Entries Per Employee By Project', self._header_format_orange)
        worksheet.write(total_row, 11, '', self._header_format_orange)
        worksheet.write(total_row, 12, '', self._header_format_orange)
        worksheet.write(total_row + 1, 10, 'Employee', self._header_format_gray)
        worksheet.write(total_row + 1, 11, 'Project', self._header_format_gray)
        worksheet.write(total_row + 1, 12, 'Hours', self._header_format_gray)

        worksheet.set_column(10,10,25,self._cell_wrap_noborder)
        worksheet.set_column(11,11,20,self._cell_wrap_noborder)
        worksheet.set_column(12,12,9,self._cell_wrap_noborder)
        worksheet.set_column(13,13,5,self._cell_wrap_noborder)

        emp_list=[]
        sub_list=[]
        add_emp= False
        weekly_hours = 0
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        for key, value in grouped_per_employee_by_project['Week Effort'].sum().items():
           
            #print(f'{g1[2]} : {v1}')
            if len(emp_list) == 0:
                emp_list.append(key[0])
                add_emp = True
            else:
                if key[0] in emp_list:
                    add_emp= False
                else:
                    emp_list.append(key[0])
                    sub_list.clear()
                    add_emp = True
                    worksheet.write(total_row + index, 10, '', self._cell_sub_total)
                    worksheet.write(total_row + index, 11, '', self._cell_sub_total)
                    worksheet.write_number(total_row + index, 12, weekly_hours, self._cell_sub_total)
                    weekly_hours = 0
                    index += 1

            if add_emp:
                worksheet.write(total_row + index, 10,f'{key[0]}', self._cell_wrap_noborder)
                worksheet.write(total_row + index, 11,f'{key[1]}', self._cell_wrap_noborder)
            else:
                worksheet.write(total_row + index, 11,f'{key[1]}', self._cell_wrap_noborder)
            
            worksheet.write_number(total_row + index, 12,value, self._cell_wrap_noborder)
            weekly_hours += value
            index += 1

        worksheet.write(total_row + index, 10, '', self._cell_sub_total)
        worksheet.write(total_row + index, 11, '', self._cell_sub_total)
        worksheet.write_number(total_row + index, 12, weekly_hours, self._cell_sub_total)
        weekly_hours = 0
        index += 1
        worksheet.write(total_row + index, 10, f'Total Hours', self._header_format_orange)
        worksheet.write(total_row + index, 11, '', self._header_format_orange)
        worksheet.write_number(total_row + index, 12, total_hrs, self._cell_total)

  
    def _load_report_settings(self) -> None:
        self._report_settings = self._config.get_section('reports')
        self._report_schedules = self._report_settings['schedules']

    def _get_report_schedule(self, month: int) -> List[str]:
        return self._report_schedules[month-1]
