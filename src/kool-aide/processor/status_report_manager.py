# kool-aide/processor/common_manager.py

import pprint
import jsonpickle
import json
from beautifultable import BeautifulTable
import pandas as pd
from tabulate import tabulate
import os


from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger
from ..library.constants import *

from ..db_access.connection import Connection
from ..db_access.dbhelper.status_report_helper import StatusReportHelper

from ..model.cli_argument import CliArgument
from ..model.aide.project import Project
from ..model.aide.week_range import WeekRange
from ..model.aide.status_report import StatusReport

from ..assets.resources.messages import *


class StatusReportManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments
        self._db_helper = StatusReportHelper(self._logger, self._config, self._connection)

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
        
        results = self._db_helper.get_status_report_view()
        data_frame = pd.DataFrame(results.fetchall()) 
        data_frame.columns = results.keys()

        columns = None
        sort_keys = None
        project_filter = None
        week_filter = None
        
        try:
            try:
                json_parameters = json.loads(arguments.parameters)
                sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                project_filter = None if PARAM_PROJECT not in json_parameters else json_parameters[PARAM_PROJECT] 
                week_filter = None if PARAM_WEEK not in json_parameters else json_parameters[PARAM_WEEK] 
                columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
            except:
                self._log(f'error reading parameters . {str(ex)}',2)
        
            if project_filter is not None and len(project_filter)>0:
                data_frame = data_frame[data_frame['Project'].isin(project_filter)]
            
            if week_filter is not None and len(week_filter)>0:
                data_frame = data_frame[data_frame['WeekRangeId'].isin(week_filter)]
            
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

            self.send_to_output(data_frame, arguments.display_format, arguments.output_file)
            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
       
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
