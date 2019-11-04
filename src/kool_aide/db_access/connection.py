# kool-aide/db_access/connection.py

from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import Select, Insert, Update, Delete
from sqlalchemy.orm import sessionmaker
import sqlalchemy as db
import pyodbc
import os
from datetime import datetime
import urllib

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger

class Connection:
    def __init__(self, config: AppSetting, logger: CustomLogger):
        self._logger = logger
        self._config = config
        self._log("creating component")
        self._connection = None
    
    def initialize(self):
        self._log("initializing")

        try:
           
            params = urllib.parse.quote_plus("DRIVER={SQL Server};"
                f"SERVER={self._config.connection_setting.server_name};"
                f"DATABASE={self._config.connection_setting.database};"
                f"UID={self._config.connection_setting.uid};"
                f"PWD={self._config.connection_setting.password}"
            )

            engine = db.create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
            
            self._connection = engine.connect()
            
            metadata = MetaData(engine)
            self.status_report_view = Table('vw_statusreport', metadata, autoload = True)
            self.week_range = Table('week_range', metadata, autoload = True)
            self.project = Table('project', metadata, autoload = True)
            self.employee = Table('employee', metadata, autoload = True)
            self.asset_inventory_view = Table('vw_assetinventory', metadata, autoload = True)
            #self.project_view = Table('vw_project', metadata, autoload = True)
            self.department = Table('department', metadata, autoload = True)
            self.division = Table('division', metadata, autoload = True)
            self.attendance = Table('attendance', metadata, autoload = True)
            self.commendation = Table('commendations', metadata, autoload = True)

            self._session = sessionmaker(bind=engine)()
            
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
        self._logger.log(f"{message} [db_access.connection]", level)

    def get_status_report_view(self, project, week):
        try:
            query = self.status_report.select(
                self.status_report.c.Project == project and 
                self.status_report.c.WeekRangeId == int(week)
            )
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False, None

    def exec(self, sp):
        self._session.execute(sp)    
    