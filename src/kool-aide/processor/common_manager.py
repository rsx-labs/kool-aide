# kool-aide/processor/common_manager.py

import pprint
import json
import pandas as pd
from tabulate import tabulate
from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger
from ..library.constants import *

from ..db_access.connection import Connection
from ..db_access.dbhelper.common_helper import CommonHelper

from ..model.cli_argument import CliArgument
from ..model.aide.project import Project
from ..model.aide.week_range import WeekRange

from ..assets.resources.messages import *


class CommonManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument = None):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments
        self._db_helper = CommonHelper(self._logger, self._config, self._connection)

        self._log("initialize")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [processor.common_manager]", level)

    def retrieve(self, arguments : CliArgument):
        self._log(f"retrieving model : {arguments.model}")
  
        if arguments.model == SUPPORTED_MODELS[3]:
            self._retrieve_project(arguments)
            return True, DATA_RETRIEVED

        if arguments.model == SUPPORTED_MODELS[2]:
            self._retrieve_week_range(arguments)
            return True, DATA_RETRIEVED
            
        return False, NOT_SUPPORTED
    
    def get_project_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_all_project()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()
            
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

    def _retrieve_project(self, arguments: CliArgument):
        
        try:
            data_frame = self.get_project_data_frame(arguments)

            self.send_to_output(data_frame, arguments.display_format, arguments.output_file)

        except Exception as ex:
            self._log(f'error retrieving data. {str(ex)}')
    
    def _retrieve_week_range(self, arguments):
        columns = None
        sort_keys = None
        week_filter = None
        try:
            results = self._db_helper.get_all_week_range()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            try:
                json_parameters = json.loads(arguments.parameters)
                sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                week_filter = None if PARAM_WEEK not in json_parameters else json_parameters[PARAM_WEEK] 
                columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS]
            except Exception as ex:
                self._log(f'error parsing papamer or parameter empty. {str(ex)}')
        
            if week_filter is not None and len(week_filter)>0:
                data_frame = data_frame[data_frame['WEEK_ID'].isin(week_filter)]
            
            if sort_keys is not None and len(sort_keys)>0:
                data_frame.sort_values(by=sort_keys, inplace= True)
        
            limit = int(arguments.result_limit)

            if columns is not None and len(columns)>0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
                
            
            if columns is not None and len(columns)>0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)

            self.send_to_output(data_frame, arguments.display_format, arguments.output_file)

        except Exception as ex:
            self._log(f'error retrieving data. {str(ex)}')

    def send_to_output(self, data_frame: pd.DataFrame, format, out_file):
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
                data_frame.to_excel(excel_file)
                print(f"the file was saved : {excel_file}") 
            elif format == DISPLAY_FORMAT[0]:    
                print('\n') 
                print(tabulate(data_frame, showindex=False, headers=data_frame.columns))
                print('\n') 
            else:
                print(NOT_SUPPORTED)
        except Exception as ex:
            self._log(str(ex),2)
