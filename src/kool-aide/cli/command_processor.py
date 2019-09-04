# kool-aide/cli/command_processor.py

from ..db_access.connection import Connection

from ..library.custom_logger import CustomLogger
from ..library.app_setting import AppSetting
from ..library.constants import *

from ..model.cli_argument import CliArgument

from ..processor.report_manager import ReportManager
from ..processor.common_manager import CommonManager
from ..processor.status_report_manager import StatusReportManager

class CommandProcessor:
    def __init__(self, logger: CustomLogger, config: AppSetting, 
        db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection
        
        self._log("creating component")

    def _log(self, message, level = 3):
        self._logger.log(f"{message} [command processor]", level)

    def delegate(self, arguments: CliArgument):
        self._log(f"delegating {str(arguments)}")

        if arguments.action == CMD_ACTIONS[0]: # create
            return True, ""
        elif arguments.action == CMD_ACTIONS[1]: # retrieve
            if arguments.model in SUPPORTED_MODELS:
                result, message = self._retrieve_model(arguments)
                return True, f"{result} | {message}"
            else:
                self._log(f"model not supported : {arguments.model}")
                return False, "model not supported"
                
        elif arguments.action == CMD_ACTIONS[2]: # update
            return True, ""
        elif arguments.action == CMD_ACTIONS[3]: # delete
            return True, ""
        elif arguments.action == CMD_ACTIONS[4]: # gen report
            if arguments.report in REPORT_TYPES:
                
                result, message = self._generate_report(arguments)

                return True, f"{result} | {message}"
            else:
                self._log(f"report type not supported : {arguments.report}")
                return False, "report type not supported"
        else:
            return False, "invalid action"

    def _generate_report(self, arguments : CliArgument):
        self._log(f"generating report : {arguments.report}")

        report_manager = ReportManager(self._logger, self._config, self._connection, arguments)
        return report_manager.generate()

    def _retrieve_model(self, arguments: CliArgument):
        self._log(f"retrieving model : {arguments.model}")

        if arguments.model == SUPPORTED_MODELS[4]:
            status_report_manager = StatusReportManager(self._logger, self._config, self._connection, arguments)
            return status_report_manager.retrieve(arguments)
        else:
            common_manager = CommonManager(self._logger, self._config, self._connection, arguments)
            return common_manager.retrieve(arguments)