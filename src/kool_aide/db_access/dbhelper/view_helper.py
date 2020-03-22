
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
                    query = self._connection.entity.select(
                        self._connection.entity.c.WeekRangeId.in_(weeks)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    
    def get_asset_inventory_view(self, departments=[], divisions=[]):
        try:
            query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_project_view(self, departments=[], divisions=[]):
        try:
            query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    
    def get_commendation_view(self):
        try:
            query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_contact_list_view(self):
        try:
            query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    
    def get_leave_summary_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_task_view(self, status=[]):
        try:
            if status is not None:
                if len(status) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.TaskStatusID.in_(status)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_action_list_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_lesson_learnt_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    
    def get_project_billability_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_employee_billability_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_concern_list_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_success_register_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_comcell_schedule_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False
    
    def get_kpi_summary_view(self, fys=[]):
        try:
            if fys is not None:
                if len(fys) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.FiscalYear.in_(fys)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_attendance_view(self, status=[]):
        try:
            if status is not None:
                if len(status) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.StatusID.in_(status)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_skills_matrix_view(self, divisions=[]):
        try:
            if divisions is not None:
                if len(divisions) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.DivisionID.in_(divisions)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False

    def get_resource_planner_view(self, months=[]):
        try:
            if months is not None:
                if len(months) > 0:
                    query = self._connection.entity.select(
                        self._connection.entity.c.MonthID.in_(months)
                    )
                else:
                    query = self._connection.entity.select()
            else:
                query = self._connection.entity.select()
            result = query.execute()
            return result
        except Exception as ex:
            self._log(f"error getting db values. {str(ex)}")
            return False


    