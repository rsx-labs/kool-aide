
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

class CommonHelper:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection

        self._log("initialize")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [db_access.dbhelper.common_helper]", level)

    def get_all_week_range(self, limit= 0):
        try:
            if limit <= 0:
                query = self._connection.week_range.select()
            else:
                query = self._connection.week_range.select().limit(limit)
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False, None



    