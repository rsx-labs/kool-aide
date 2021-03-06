# kool-aide/processors/attendance_manager.py

import json
import pandas as pd
from tabulate import tabulate
from datetime import date, datetime
import os


from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import print_to_screen, get_end_date, \
    get_start_date, get_param_value

from kool_aide.db_access.connection import Connection
from kool_aide.db_access.dbhelper.attendance_helper import AttendanceHelper

from kool_aide.model.cli_argument import CliArgument
from kool_aide.model.aide.project import Project

from kool_aide.assets.resources.messages import *

class AttendanceManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments
        self._db_helper = AttendanceHelper(
            self._logger, 
            self._config, 
            self._connection
        )

        self._log("creating component")

    def _log(self, message, level=3) -> None:
        self._logger.log(f"{message} [processor.attendance_manager]", level)

    
    def create(self, arguments: CliArgument) ->(bool, str):
        pass

    def retrieve(self, arguments: CliArgument)->(bool, str):
        self._log(f"retrieving model : {arguments.model}")
  
        if arguments.model == SUPPORTED_MODELS[1]:
            self._retrieve(arguments)
            return True, DATA_RETRIEVED

        return False, NOT_SUPPORTED

    def update(self, arguments: CliArgument) ->(bool, str):
        pass

    def delete(self, arguments: CliArgument) ->(bool, str):
        pass

    def execute(self, arguments: CliArgument) ->(bool, str):
        pass

    
    def get_data_frame(self, arguments : CliArgument) -> pd.DataFrame:
        columns = None
        sort_keys = None
        ids = None 
        start_date = None
        end_date = None
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters) 
                    columns = get_param_value(PARAM_COLUMNS, json_parameters) 
                    ids = get_param_value(PARAM_IDS, json_parameters) 
                    
                    temp_date = get_param_value(PARAM_START_DATE, json_parameters)       
                    start_date = date.today() if temp_date is None \
                    else datetime.strptime(temp_date,'%m-%d-%Y')
                    
                    temp_date = get_param_value(PARAM_END_DATE, json_parameters)       
                    end_date = date.today() if temp_date is None \
                    else datetime.strptime(temp_date,'%m-%d-%Y')

                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_all_attendance()
            data_frame = pd.DataFrame(results.fetchall())
            data_frame.columns = results.keys()
            
            if sort_keys is not None and len(sort_keys)>0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['EMP_ID'].isin(ids)]

            if start_date is not None and end_date is not None:
                data_frame = data_frame[
                    (data_frame['DATE_ENTRY']>= get_start_date(start_date)) &
                    (data_frame['DATE_ENTRY']<= get_end_date(end_date))
                ]

            limit = int(arguments.result_limit)

            if columns is not None and len(columns)>0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
            
            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
            return None

    def send_to_output(self, data_frame: pd.DataFrame, format, out_file)-> None:
        if out_file is None:
            file = DEFAULT_FILENAME
            out_file = file

        out_file = append_date_to_file_name(out_file)
        try:
            if format == OUTPUT_FORMAT[1]:
                json_file = f"{file}.json" if out_file is None else out_file
                data_frame.to_json(json_file, orient='records')
                # print(f"the file was saved : {json_file}")
            elif format == OUTPUT_FORMAT[2]:
                csv_file = f"{file}.csv" if out_file is None else out_file
                data_frame.to_csv(csv_file)
                # print(f"the file was saved : {csv_file}")
            elif format == OUTPUT_FORMAT[3]:
                excel_file = f"{file}.xslx" if out_file is None else out_file
                data_frame.to_excel(excel_file)
                # print(f"the file was saved : {excel_file}") 
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

    def _retrieve(self, arguments: CliArgument) -> None:
        try:
            data_frame = self.get_data_frame(arguments)

            self.send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file
            )

        except Exception as ex:
            self._log(f'error retrieving data. {str(ex)}')

