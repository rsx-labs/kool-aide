# kool-aide/processor/common_manager.py

import pprint
import jsonpickle
import json

from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger
from ..library.constants import *

from ..db_access.connection import Connection
from ..db_access.common_dbhelper import CommonDBHelper

from ..model.cli_argument import CliArgument
from ..model.aide_project import AIDEProject
from ..model.aide_week_range import AIDEWeekRange


class CommonManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments
        self._db_helper = CommonDBHelper(self._logger, self._config, self._connection)

        self._log("initialize")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [common manager]", level)

    def retrieve(self, arguments : CliArgument):
        self._log(f"retrieving model : {arguments.model}")
  
        if arguments.model == SUPPORTED_MODELS[3]:
            self._retrieve_project(arguments)

        if arguments.model == SUPPORTED_MODELS[2]:
            self._retrieve_week_range(arguments)
            
        return True, "retrieved data"
    
    def _retrieve_project(self, arguments, parameters = ""):
        results = self._db_helper.get_all_project()
        projects = []
        for result in results:
            project = AIDEProject(result)
            projects.append(self._format_result(project, arguments.is_csv_format))
            
        self.print_to_screen(projects, arguments.is_csv_format)
    
    def _retrieve_week_range(self, arguments, parameters = ""):
        results = self._db_helper.get_all_week_range()
        week_ranges = []
        for result in results:
            week = AIDEWeekRange(result)
            week_ranges.append(self._format_result(week, arguments.is_csv_format))
            
        self.print_to_screen(week_ranges, arguments.is_csv_format)

    def print_to_screen(self, data, is_csv = False):
        if not is_csv:
            print(json.dumps(data, indent=4, sort_keys=True))
        else:
            print(list(data))

    def _format_result(self, data, is_csv= False ):
        if not is_csv:
            return data.to_json()
        else:
            return data.to_csv()