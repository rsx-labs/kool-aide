# kool-aide/processors/report_manager.py

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
        # result = self._connection.get_status_report_view("IBP",2)
        # for stat in result:
        #     self._log(str(stat))
    
    def write_to_file(def, data):
        self._log(f"writing to {self._arguments.output_file}")

    