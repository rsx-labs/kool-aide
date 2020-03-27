import json
import pandas as pd
from tabulate import tabulate
from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import *

from kool_aide.db_access.connection import Connection
from kool_aide.db_access.dbhelper.division_helper import DivisionHelper

from kool_aide.model.cli_argument import CliArgument
from kool_aide.model.aide.project import Project

from kool_aide.assets.resources.messages import *


class DivisionManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument = None):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments
        self._db_helper = DivisionHelper(
            self._logger, 
            self._config, 
            self._connection
        )

        self._log("creating component")

    def _log(self, message, level=3) -> None:
        self._logger.log(f"{message} [processor.division_manager]", level)

    def create(self, arguments: CliArgument)->(bool, str):
        if arguments.input_file is None:
            return False, MISSING_PARAMETER.replace('%0', 'input file')

        if arguments.display_format is None:
            # default to json
            arguments.display_format = OUTPUT_FORMAT[0]

        if arguments.display_format == OUTPUT_FORMAT[0]: #json
            divisions = self._read_json(arguments.input_file)
            self._log(f'read data = {divisions}')

            for division in divisions:
                result, error = self._add(division, arguments)
                self._log(f'inserting division = {result} ; \
                            error = {error}')
            
            print_to_screen(f'Done creating division/s. Check the logs.', arguments.quiet_mode)
            return True, ''
        else:
            return False, NOT_SUPPORTED

    def _read_json(self, file: str):
        divisions = []
        with open(file) as input_file:
            divisions = json.load(input_file)
        return divisions

    def retrieve(self, arguments: CliArgument)->(bool, str):
        self._log(f"retrieving model : {arguments.model}")
  
        if arguments.model == SUPPORTED_MODELS[5]:
            self._retrieve(arguments)
            return True, DATA_RETRIEVED
        else:
            return False, NOT_SUPPORTED
    
    def update(self, arguments: CliArgument)->(bool, str):
        pass

    def delete(self, arguments: CliArgument)->(bool, str):
        pass

    def execute(self, arguments: CliArgument) -> (bool, str):
        pass

    def get_data_frame(self, arguments : CliArgument) -> pd.DataFrame:
        columns = None
        sort_keys = None
        ids = None 
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                    ids = None if PARAM_IDS not in json_parameters else json_parameters[PARAM_IDS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()
            
            if sort_keys is not None and len(sort_keys)>0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['DIV_ID'].isin(ids)]
                
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

    def _add(self, division_data, argument: CliArgument):
        
        self._log(f"division [{division_data['DIV_ID']}]")
        if division_data['DIV_ID'] is not None:
            result, error = self._db_helper.insert(
                division_data['DIV_ID'], 
                division_data['DESCR']
            )
            if result:
                print_to_screen(f'Successful inserting {division_data["DIV_ID"]}', argument.quiet_mode)
                return True, ''
            else:
                print_to_screen(f'Failed inserting {division_data["DIV_ID"]}', argument.quiet_mode)
                return False, error      
        else:
            return False, MISSING_PARAMETER

    