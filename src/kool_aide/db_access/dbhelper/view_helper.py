
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import Select, Insert, Update, Delete
import sqlalchemy as db
import pyodbc
from datetime import datetime
import urllib
import pandas as pd

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.db_access.connection import Connection

class ViewHelper:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection

        self._log("initialize")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [db_access.dbhelper.view_helper]", level)

    def get_status_report_view(self, weeks=[]):
        try:
            if weeks is not None:
                if len(weeks) > 0:
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
    
    def get_asset_inventory_view(self, departments=[], divisions=[]):
        try:
            query = self._connection.asset_inventory_view.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_project_view(self, departments=[], divisions=[]):
        try:
            query = self._connection.project_view.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    
    def get_commendation_view(self):
        try:
            query = self._connection.commendation_view.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_contact_list_view(self):
        try:
            query = self._connection.contact_view.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    
    def get_leave_summary_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.leave_sumarry_view.select(
                        self._connection.leave_sumarry_view.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.leave_sumarry_view.select()
            else:
                query = self._connection.leave_sumarry_view.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    