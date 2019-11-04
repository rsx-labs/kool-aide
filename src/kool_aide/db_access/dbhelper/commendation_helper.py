
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import Select, Insert, Update, Delete
import sqlalchemy as db
import pyodbc
import os
from datetime import datetime
import urllib

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.db_access.connection import Connection

class CommendationHelper:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection

        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [db_access.dbhelper.commendation_helper]", level)

    def get(self, ids=[]):
        try:
            query = self._connection.commendation.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return None
    
    def insert(self, employee, project, date_sent, 
                sent_by, reason, emp_id) -> bool:
        try:
            command = self._connection.commendation.insert().values(
                EMPLOYEE = employee,
                PROJECT = project,
                DATE_SENT = date_sent,
                SENT_BY = sent_by,
                REASON = reason,
                EMP_ID = emp_id,
                DATE = datetime.now()
            )
            result = command.execute()
            self._logger.log(result, 4)
            return True, ''
        except Exception as ex:
            self._logger.log("error inserting data to db...", 1)
            self._logger.log(f"error : {ex}", 1)
            return False, str(ex)
    
    
