# kool-aide/cli/command_processor.py

from ..db_access.connection import Connection
from ..library.custom_logger import CustomLogger
from ..library.app_setting import AppSetting
from ..library.constants import *
from ..model.cli_argument import CliArgument

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
            return True, ""
        elif arguments.action == CMD_ACTIONS[2]: # update
            return True, ""
        elif arguments.action == CMD_ACTIONS[3]: # delete
            return True, ""
        elif arguments.action == CMD_ACTIONS[4]: # gen report
            if arguments.report in REPORT_TYPES:
                self._log(f"generating report : {arguments.report}")
                return True, ""
            else:
                self._log(f"report type not supported : {arguments.report}")
                return False, "report type not supported"
        else:
            return False, "invalid action"