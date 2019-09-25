
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

class EmployeeHelper:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection

        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [db_access.dbhelper.employee_helper]", level)

    def get_all_employee(self, ids=[]):
        try:
            if ids is not None and len(ids) > 0:
                query = self._connection.employee.select(
                    self._connection.employee.c.WeekRangeId.in_(ids)
                )
            else:
                query = self._connection.employee.select()
           
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    
    def insert(self, user_session_id, start_credit, starting_demo_credit):
        try:
            if self._db_connection.closed:
                self._logger.log("db connection closed, reconnecting ...")
                self._initialize()

            command = self._trade_session_table.insert().values(
                client_session_id = user_session_id,
                start_timestamp  = datetime.now(),
                start_credit = start_credit,
                start_demo_credit = starting_demo_credit
            )
            result = self._db_connection.execute(command)
            # self._logger.log(result, 4)
            return result.lastrowid
        except Exception as ex:
            self._logger.log("error inserting data to db...", 1)
            self._logger.log(f"error : {ex}", 1)
            return None
