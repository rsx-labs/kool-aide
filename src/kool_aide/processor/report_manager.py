# kool-aide/processors/report_manager.py
import pandas as pd

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.db_access.connection import Connection
from kool_aide.model.cli_argument import CliArgument
from kool_aide.processor.status_report_manager import StatusReportManager
from kool_aide.processor.common_manager import CommonManager

class ReportManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                    db_connection: Connection, arguments: CliArgument):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments

        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [processor.report_manager]", level)

    def generate(self, arguments: CliArgument):
        self._log("report generation started")
        # get target report from argument
        if self._arguments.report == REPORT_TYPES[0]:
            # weekly status
            self._generate_status_report(arguments)
            return True, ''
        else:
            pass

    def _write_to_file(self, data):
        self._log(f"writing to {self._arguments.output_file}")

    def _generate_status_report(self, arguments: CliArgument):
        # for now just call status report manager, 
        # TODO: Move all report genaration code here, status report manager will only retrieve
        status_report_manager = StatusReportManager(
            self._logger, 
            self._config, 
            self._connection
        )

        return status_report_manager.retrieve(arguments)
      
       
        



    