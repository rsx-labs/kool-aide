
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import Select, Insert, Update, Delete
import sqlalchemy as db
import pyodbc
from datetime import datetime
import urllib

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.db_access.connection import Connection

class WeekRangeHelper:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection

        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [db_access.dbhelper.week_range_helper]", level)

    def get(self, ids=[]):
        try:
            query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return None
    
    def insert(self, start, end) -> bool:
        try:
            command = self._connection.entity.insert().values(
                WEEK_START = start,
                WEEK_END = end  # for now skip this
            )
            result = command.execute()
            self._logger.log(result, 4)
            return True, ''
        except Exception as ex:
            self._logger.log("error inserting data to db...", 1)
            self._logger.log(f"error : {ex}", 1)
            return False, str(ex)
    
    
