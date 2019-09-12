# kool-aide/processors/attendance_manager.py

import pprint
import jsonpickle
import json
from beautifultable import BeautifulTable
import pandas as pd
from tabulate import tabulate
import os


from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *

from kool_aide.db_access.connection import Connection
from kool_aide.db_access.dbhelper.status_report_helper import StatusReportHelper

from kool_aide.model.cli_argument import CliArgument
from kool_aide.model.aide.project import Project
from kool_aide.model.aide.week_range import WeekRange
from kool_aide.model.aide.status_report import StatusReport

from kool_aide.assets.resources.messages import *

class AttendanceManager:
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

    def record_time_in(self, user, password = ""):
        self._connection.exec('sp_GetAllDepartment')