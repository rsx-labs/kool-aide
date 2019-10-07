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

class StatusReport:

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
        self._logger.log(f"{message} [report.status_report]", level)

    def generate(self, format: str) -> None:
        if format == OUTPUT_FORMAT[3]: # excel
            return self._generate_excel_report()
        else:
            return False, NOT_SUPPORTED

    def _generate_excel_report(self) -> None:
        try:
            data_frame = self._data
            self._workbook = self._writer.book
            self._main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            self._header_format_orange = self._workbook.add_format(SHEET_HEADER_ORANGE)
            self._header_format_gray = self._workbook.add_format(SHEET_HEADER_GRAY)
            self._cell_wrap_noborder = self._workbook.add_format(SHEET_CELL_WRAP_NOBORDER)
            self._cell_total = self._workbook.add_format(SHEET_HEADER_LT_GREEN)
            self._cell_sub_total = self._workbook.add_format(SHEET_HEADER_GAINSBORO)

            drop_columns = ['WeekRangeStart', 'WeekRangeId', 'ProjectId']
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
                'DivisionID'
            ]

            data_frame.drop(drop_columns, inplace=True, axis=1)
            data_frame.columns = column_headers

            self._create_main_report(data_frame)
            self._create_summary_by_week_report(data_frame)
            self._create_summary_report(data_frame)
                          
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            # print(f"the file was saved : {file_name}") 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_main_report(self, data_frame: pd.DataFrame)-> None:
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
        worksheet.write(total_row + 1, 0, 'Project', self._header_format_gray)
        worksheet.write(total_row + 1, 1, 'Hours', self._header_format_gray)

        worksheet.set_column(0,0,20,self._cell_wrap_noborder)
        worksheet.set_column(1,1,9,self._cell_wrap_noborder)
        worksheet.set_column(2,2,5,self._cell_wrap_noborder)
        project_list=[]
        sub_list = []
        index = 2
        total_hrs = sum(data_frame['Week Effort'])
        count = 0
        for key, value in grouped_per_project['Week Effort'].sum().items():
            worksheet.write(total_row + index, 0,f'{key}', self._cell_wrap_noborder)
            worksheet.write_number(total_row + index, 1,value, self._cell_wrap_noborder)
            index += 1
            count += 1
        
        worksheet.write(total_row + index, 0, f'Total Hours', self._header_format_orange)
        worksheet.write_number(total_row + index, 1, total_hrs, self._cell_total)

        # # Create a chart object.
        # summary_project_hours_chart = self._workbook.add_chart({'type': 'pie'})

        # # Configure the chart from the dataframe data. Configuring the segment
        # # colours is optional. Without the 'points' option you will get Excel's
        # # default colours.
        # summary_project_hours_chart.add_series({
        #     'categories': "='Team Summary'!A3:A13",
        #     'values':     "='Team Summary'!B3:B13",
        # })

        # # Insert the chart into the worksheet.
        # worksheet.insert_chart('A16', summary_project_hours_chart)


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

  
    


    
