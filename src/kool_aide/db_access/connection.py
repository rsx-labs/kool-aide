# kool-aide/db_access/connection.py

from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import Select, Insert, Update, Delete
from sqlalchemy.orm import sessionmaker
import sqlalchemy as db
import pyodbc
# import os
from datetime import datetime
import urllib

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import MAP_MODEL_TO_DB, MAP_VIEW_TO_DB

class Connection:
    def __init__(self, config: AppSetting, logger: CustomLogger):
        self._logger = logger
        self._config = config
        self._log("creating component")
        self._connection = None
    
    def initialize(self, name = '', is_view = None):
        self._log("initializing")

        try:
            self._log("building connection string ...", 4)
            params = urllib.parse.quote_plus("DRIVER={SQL Server};"
                f"SERVER={self._config.connection_setting.server_name};"
                f"DATABASE={self._config.connection_setting.database};"
                f"UID={self._config.connection_setting.uid};"
                f"PWD={self._config.connection_setting.password}"
            )
            self._log("creating engine ...", 4)
            engine = db.create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
            self._log("engine connecting ...", 4)
            self._connection = engine.connect()
            
            metadata = MetaData(engine)
            
            self._log("instantiating tables and views ...", 4)
            
            self.entity = None

            if is_view is not None:
                db_entity = ''
                if not is_view:  
                    # the entity is a model
                    db_entity = MAP_MODEL_TO_DB[name]
                else:
                    # is a view
                    db_entity = MAP_VIEW_TO_DB[name]

                self.entity = Table(db_entity, metadata, autoload = True)
            
            self._log("creating session ...", 4)
            self._session = sessionmaker(bind=engine)()
            self._log("initialization done ...", 4)
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
    