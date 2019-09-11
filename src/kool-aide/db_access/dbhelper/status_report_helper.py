
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

class StatusReportHelper:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection

        self._log("initialize")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [db_access.dbhelper.status_report_helper]", level)

    def get_status_report_view(self, weeks=[]):
        try:
            if weeks is not None:
                if len(weeks)>0:
                    query = self._connection.status_report_view.select(
                        self._connection.status_report_view.c.WeekRangeId.in_(weeks)
                    )
                else:
                    query = self._connection.status_report_view.select()
            else:
                query = self._connection.status_report_view.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
        
    # def get_all_week_range(self):
    #     try:
    #         query = self._connection.week_range.select()
    #         result = query.execute()
    #         return result
    #     except Exception as ex:
    #         self._log(f"error getting db values. {str(ex)}")
    #         return False, None

    # def get_all_project(self, limit=0):
    #     try:
    #         if limit <= 0:
    #             query = self._connection.project.select()
    #         else:
    #             query = self._connection.project.select().limit(limit)
    #         result = query.execute()
    #         return result
    #     except Exception as ex:
    #         self._log(f"error getting db values. {str(ex)}")
    #         return False, None
    
