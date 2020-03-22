
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import Select, Insert, Update, Delete
import sqlalchemy as db
import pyodbc
from datetime import datetime
import urllib

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.db_access.connection import Connection

from kool_aide.model.aide.employee import Employee

class EmployeeHelper:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection):
        self._logger = logger
        self._config = config
        self._connection = db_connection

        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [db_access.dbhelper.employee_helper]", level)

    def get(self, ids=[]):
        try:
            if ids is not None and len(ids) > 0:
                query = self._connection.entity.select(
                    self._connection.entity.c.WeekRangeId.in_(ids)
                )
            else:
                query = self._connection.entity.select()
           
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return None
    
    def insert(self, employee: Employee) -> bool:
        try:
            command = self._connection.entity.insert().values(
                EMP_ID = employee.id,
                WS_EMP_ID = employee.custom_id,
                LAST_NAME = employee.last_name,
                FIRST_NAME = employee.first_name,
                MIDDLE_NAME = employee.middle_name,
                NICK_NAME = employee.nick_name,
                BIRTHDATE = employee.birth_date,
                POS_ID = employee.position_id,
                DATE_HIRED = employee.hire_date,
                STATUS = employee.status,
                IMAGE_PATH = employee.image_path,
                GRP_ID = employee.group_id,
                DEPT_ID = employee.department_id,
                ACTIVE = employee.is_active,
                DIV_ID = employee.division_id,
                SHIFT_STATUS = employee.status,
                APPROVED = employee.is_approved
            )
            result = command.execute()
            self._logger.log(result, 4)
            return True, ''
        except Exception as ex:
            self._logger.log("error inserting data to db...", 1)
            self._logger.log(f"error : {ex}", 1)
            return False, str(ex)
