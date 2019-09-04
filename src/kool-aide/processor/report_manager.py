# kool-aide/processors/report_manager.py
import pandas as pd

from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger
from ..library.constants import *
from ..db_access.connection import Connection
from ..model.cli_argument import CliArgument


class ReportManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments

        self._log("initialize")

    def _log(message, level=3):
        self._logger.log(f"{message} [report manager]", level)

    def generate(self):
        self._log("report generation started")
        # get target report from argument
        if self._arguments.report == REPORT_TYPES[0]:
            # team weekly status
            pass
        else:
            pass

    def _write_to_file(self, data):
        self._log(f"writing to {self._arguments.output_file}")

    def _generate_tl_monthly_report(self):
        pass

    