# kool-aide/processor/common_manager.py

import pandas as pd
from tabulate import tabulate
from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import *

from kool_aide.db_access.connection import Connection
from kool_aide.db_access.dbhelper.common_helper import CommonHelper
from kool_aide.db_access.dbhelper.week_range_helper import WeekRangeHelper

from kool_aide.model.cli_argument import CliArgument

from kool_aide.assets.resources.messages import *


class CommonManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument = None):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments
        self._common_db_helper = CommonHelper(self._logger, self._config, self._connection)
        self._week_range_db_helper = WeekRangeHelper(self._logger, self._config, self._connection)
        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [processor.common_manager]", level)

    def retrieve(self, arguments : CliArgument):
        self._log(f"retrieving model : {arguments.model}")
  
        if arguments.model == SUPPORTED_MODELS[2]:
            self._retrieve_week_range(arguments)
            return True, DATA_RETRIEVED
            
        return False, NOT_SUPPORTED

    def create(self, arguments: CliArgument)->(bool, str):
        if arguments.input_file is None:
            return False, MISSING_PARAMETER.replace('%0', 'input file')

        if arguments.display_format is None:
            # default to json
            arguments.display_format = OUTPUT_FORMAT[0]

        if arguments.display_format == OUTPUT_FORMAT[0]: #json
            elements = self._read_json(arguments.input_file)
            self._log(f'read data = {elements}')

            for element in elements:
                result, error = self._add(element, arguments)
                self._log(f'inserting data = {result} ; \
                            error = {error}')
            
            print_to_screen(f'Done creating data. Check the logs.', arguments.quiet_mode)
            return True, ''
        else:
            return False, NOT_SUPPORTED

    def _read_json(self, file: str):
        elements = []
        with open(file) as input_file:
            elements = json.load(input_file)
        return elements
      
    def _retrieve_week_range(self, arguments):
        columns = None
        sort_keys = None
        week_filter = None
        try:
            results = self._week_range_db_helper.get()
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
            if format == OUTPUT_FORMAT[1]:
                json_file = f"{file}.json" if out_file is None else out_file
                data_frame.to_json(json_file, orient='records')
                print(f"the file was saved : {json_file}")
            elif format == OUTPUT_FORMAT[2]:
                csv_file = f"{file}.csv" if out_file is None else out_file
                data_frame.to_csv(csv_file)
                print(f"the file was saved : {csv_file}")
            elif format == OUTPUT_FORMAT[3]:
                excel_file = f"{file}.xslx" if out_file is None else out_file
                data_frame.to_excel(excel_file)
                print(f"the file was saved : {excel_file}") 
            elif format == OUTPUT_FORMAT[0]:    
                print('\n') 
                print(tabulate(data_frame, showindex=False, headers=data_frame.columns))
                print('\n') 
            else:
                print(NOT_SUPPORTED)
        except Exception as ex:
            self._log(str(ex),2)

    def _add(self, element, arguments):

        if arguments.model == SUPPORTED_MODELS[2]: #week range
            # self._log(f'adding week_range : {element["WEEK_ID"]}')
            if True:
                result, error = self._week_range_db_helper.insert(
                    get_date(element['WEEK_START']),
                    get_date(element['WEEK_END'])
                )
                element_id_string = f' week range' # id :{element["WEEK_ID"]}'
            else:
                result = False
                error = MISSING_PARAMETER
        else:
            result = False
            error = NOT_SUPPORTED
        
        if result:
            print_to_screen(f'Successful inserting {element_id_string}', arguments.quiet_mode)
            return True, ''
        else:
            print_to_screen(f'Failed inserting {element_id_string}', arguments.quiet_mode)
            return False, error   

