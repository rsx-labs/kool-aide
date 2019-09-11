
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import Select, Insert, Update, Delete
import sqlalchemy as db
import pyodbc
import os
from datetime import datetime
import urllib
import pandas as pd

from ...library.app_setting import AppSetting
from ...library.custom_logger import CustomLogger
from ...db_access.connection import Connection

class AttendancetHelper:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection

        self._log("initialize")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [db_access.dbhelper.attendance_helper]", level)

    # def get_status_report_view(self, projects=[], weeks=[]):
    #     try:
    #         query = self._connection.status_report_view.select()
    #         result = query.execute()
    #         return result
    #     except Exception as ex:
    #         self._log(f"error getting db values. {str(ex)}")
    #         return False
    def record_time_in(self, user, password=""):
        try:
            
        except Exception as ex:
            self._log(f"error recording time-in. {str(ex)}")
            return False
        
    
