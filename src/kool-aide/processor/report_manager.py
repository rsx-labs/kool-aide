# kool-aide/processors/report_manager.py
import pandas as pd

from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger
from ..library.constants import *
from ..db_access.connection import Connection
from ..model.cli_argument import CliArgument
from .status_report_manager import StatusReportManager
from .common_manager import CommonManager

class ReportManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments

        self._log("initialize")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report manager]", level)

    def generate(self, arguments: CliArgument):
        self._log("report generation started")
        # get target report from argument
        if self._arguments.report == REPORT_TYPES[0]:
            # weekly status
            self._generate_weekly_report(arguments)
            return True, ''
        else:
            pass

    def _write_to_file(self, data):
        self._log(f"writing to {self._arguments.output_file}")

    def _generate_weekly_report(self, arguments: CliArgument):
        status_report_manager = StatusReportManager(self._logger, self._config, self._connection)
        project_manager = CommonManager(self._logger, self._config, self._connection)
        df_status = status_report_manager.get_data_frame(arguments)
        df_project = project_manager.get_project_data_frame(arguments)
        df_project = df_project[df_project['DSPLY_FLG'] == 1]
        print(df_project)



    