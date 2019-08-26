# kool-aide/db_access/connection.py

from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import Select, Insert, Update, Delete
import sqlalchemy as db
import pyodbc
import os
from datetime import datetime
import urllib

from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger

class Connection:
    def __init__(self, config: AppSetting, logger: CustomLogger):
        self._logger = logger
        self._config = config
        self._log("creating component")
        self._connection = None
    
    def initialize(self):
        self._log("initializing")

        try:
            # engine = db.create_engine('mssql+pyodbc://aide:aide1234@192.168.1.39\\SQLEXPRESS2014:1433/AIDE?driver=SQL+Server')
            # self._log("database connection is initialized")
            #self._connection = engine.connect()
            params = urllib.parse.quote_plus("DRIVER={SQL Server};"
                f"SERVER={self._config.connection_setting.server_name};"
                f"DATABASE={self._config.connection_setting.database};"
                f"UID={self._config.connection_setting.uid};"
                f"PWD={self._config.connection_setting.password}"
            )

            engine = db.create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
            
            self._connection = engine.connect()
            
            metadata = MetaData(engine)
            self._status_report = Table('vw_statusreport', metadata, autoload = True)
            self._week_range = Table('week_range', metadata, autoload = True)
            #self._log(week.columns.key())
            
            # query = week.select()
            # res = self._connection.execute(query)
            # for result in res:
            #     self._log(f"{str(result)}")
            return True   
        except Exception as ex:
            self._log(f"error initializing db. {str(ex)}", 1)
            return False

        return True
        
    def _log(self, message, level = 3):
        self._logger.log(f"{message} [connection]", level)

    def get_status_report_view(self, project, week):
        try:
            query = self._status_report.select(
                self._status_report.c.Project == project and 
                self._status_report.c.WeekRangeId == int(week)
            )
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False, None