# kool-aide/processors/attendance_manager.py

import pprint
import jsonpickle
import json
from beautifultable import BeautifulTable
import pandas as pd
from tabulate import tabulate
import os


from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger
from ..library.constants import *

from ..db_access.connection import Connection
from ..db_access.dbhelper.status_report_helper import StatusReportHelper

from ..model.cli_argument import CliArgument
from ..model.aide.project import Project
from ..model.aide.week_range import WeekRange
from ..model.aide.status_report import StatusReport

from ..assets.resources.messages import *

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