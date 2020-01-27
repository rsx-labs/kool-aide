# kool-aide/processor/common_manager.py

import json
import pandas as pd
from tabulate import tabulate
import os
from datetime import datetime
import xlsxwriter
from typing import List
from datetime import datetime

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import append_date_to_file_name, \
    get_param_value, get_version, get_cell_range_address, get_cell_address

from kool_aide.db_access.connection import Connection
from kool_aide.db_access.dbhelper.view_helper \
    import ViewHelper

from kool_aide.model.cli_argument import CliArgument
from kool_aide.model.aide.project import Project

from kool_aide.assets.resources.messages import *


class ViewManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument = None):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments

        self._db_helper = ViewHelper(
            self._logger, 
            self._config, 
            self._connection
        )

        self._report_settings = None
        self._report_schedules = None
        self._report_defaults = None

        self._load_report_settings()

        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [processor.view_manager]", level)

    def retrieve(self, arguments : CliArgument):
        self._log(f"retrieving view : {arguments.model}")
  
        if arguments.view == SUPPORTED_VIEWS[0]:
            self._retrieve_status_report_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[1]:
            self._retrieve_asset_inventory_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[2]:
            self._retrieve_commendation_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[3]:
            self._retrieve_contact_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[4]:
            self._retrieve_leave_summary_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[5]:
            self._retrieve_task_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[6]:
            self._retrieve_action_list_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[7]:
            self._retrieve_lesson_learnt_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[8]:
            self._retrieve_project_billability_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[9]:
            self._retrieve_employee_billability_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[10]:
            self._retrieve_concern_list_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[11]:
            self._retrieve_success_register_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[12]:
            self._retrieve_comcell_schedule_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[13]:
            self._retrieve_kpi_summary_view(arguments)
            return True, DATA_RETRIEVED
        elif arguments.view == SUPPORTED_VIEWS[14]:
            self._retrieve_attendance_view(arguments)
            return True, DATA_RETRIEVED
        
       
        return False, NOT_SUPPORTED

    def get_status_report_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        project_filter = None
        week_filter = None
        division_filter = None
        department_filter = None

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    project_filter = None if PARAM_PROJECT not in json_parameters else json_parameters[PARAM_PROJECT] 
                    week_filter = None if PARAM_WEEK not in json_parameters else json_parameters[PARAM_WEEK] 
                    division_filter = None if PARAM_DIVISIONS not in json_parameters else json_parameters[PARAM_DIVISIONS] 
                    department_filter = None if PARAM_DEPARTMENTS not in json_parameters else json_parameters[PARAM_DEPARTMENTS] 
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            # if running as auto, use settings. else, use params
            if arguments.auto_mode:
                week_filter = self._get_report_schedule(datetime.now().month)
           
            results = self._db_helper.get_status_report_view(week_filter)
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if project_filter is not None and len(project_filter)>0:
                data_frame = data_frame[data_frame['Project'].isin(project_filter)]

            if division_filter is not None and len(division_filter)>0:
                data_frame = data_frame[data_frame['DivisionID'].isin(division_filter)]
            
            if department_filter is not None and len(department_filter)>0:
                data_frame = data_frame[data_frame['DepartmentID'].isin(department_filter)]

            if sort_keys is not None and len(sort_keys)>0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns)>0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
            
            data_frame = data_frame[data_frame['ActualWeekWork'] > 0]

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
            return None
    
    def get_asset_inventory_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        divisions = None
        departments = None
        
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    divisions = None if PARAM_DIVISIONS not in json_parameters else json_parameters[PARAM_DIVISIONS] 
                    departments = None if PARAM_DEPARTMENTS not in json_parameters else json_parameters[PARAM_DEPARTMENTS] 
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_asset_inventory_view()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if departments is not None and len(departments)>0:
                data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns) > 0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
            
            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_project_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        divisions = None
        departments = None
        projects = None
        
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = None if PARAM_SORT not in json_parameters else json_parameters[PARAM_SORT]
                    divisions = None if PARAM_DIVISIONS not in json_parameters else json_parameters[PARAM_DIVISIONS] 
                    departments = None if PARAM_DEPARTMENTS not in json_parameters else json_parameters[PARAM_DEPARTMENTS] 
                    projects = None if PARAM_PROJECT not in json_parameters else json_parameters[PARAM_PROJECT] 
                    columns = None if PARAM_COLUMNS not in json_parameters else json_parameters[PARAM_COLUMNS] 
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_project_view()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if projects is not None and len(projects)>0:
                data_frame = data_frame[data_frame['ProjectName'].isin(projects)]

            if departments is not None and len(departments)>0:
                data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns) > 0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame = data_frame.head(limit)
            
            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_commendation_view_data_frame(self, arguments : CliArgument):
        
        columns = None
        sort_keys = None
        months = None
        year = None
        projects = None
        fys = None

        col_names =[
                'Employee', 'Project', 'Sent By', 'Date Sent',
                'Month', 'Year', 'Reason', 'Fiscal Year'
            ]
        
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    months = get_param_value(PARAM_MONTHS, json_parameters, months)
                    year = get_param_value(PARAM_YEAR, json_parameters, year)
                    projects = get_param_value(PARAM_PROJECT, json_parameters)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_commendation_view()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            # data_frame['Month'] =data_frame['DateSent'].dt.month
            if arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = None
                year = None

            if projects is not None and len(projects)>0:
                data_frame = data_frame[data_frame['Project'].isin(projects)]

            if year is not None:
                data_frame = data_frame[data_frame['Year'] == year]
            
            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['Month'].isin(months)]

            if fys is not None and len(fys)>0:
                data_frame = data_frame[data_frame['FiscalYear'].isin(fys)]

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns) > 0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                data_frame.drop(['Month','Year'], inplace=True, axis=1)
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_contact_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        ids = None
        departments = None
        divisions = None
        isActive = None
        col_names =[
                'Employee ID', 'Employee Name', 'Local No', 'Mobile No',
                'Home Phone', 'Other Phone', 'Office Email',
                'Other Email', 'IsActive','DepartmentID', 'DivisionID',
                'Department', 'Division'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    ids = get_param_value(PARAM_IDS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    isActive = get_param_value(PARAM_FLAG, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_contact_list_view()
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['EmployeeID'].isin(ids)]

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if isActive is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(isActive)]
                except:
                    pass

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns) > 0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                data_frame.drop(['DepartmentID','DivisionID','IsActive'], inplace=True, axis=1)
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_leave_summary_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        ids = None
        departments = None
        divisions = None
        isActive = None
        fys = None
        types = None
        col_names =[
                'Employee ID', 'Employee Name', 'Leave Type', 'Total Leaves',
                'Used Leaves', 'Remaining Leaves', 'Remaining Mandatory Leaves',
                'Fiscal Year', 'DivisionID','DepartmentID','IsActive','LeaveType'
            ]
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    ids = get_param_value(PARAM_IDS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    isActive = get_param_value(PARAM_FLAG, json_parameters)
                    fys= get_param_value(PARAM_FYS, json_parameters)
                    types = get_param_value(PARAM_TYPES, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            

            if arguments.auto_mode:
                fys = [self._report_defaults['fiscal_year']]

            results = self._db_helper.get_leave_summary_view(fys)
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['EmployeeID'].isin(ids)]

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if types is not None and len(types)>0:
                  data_frame = data_frame[data_frame['LeaveType'].isin(types)]

            if isActive is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(isActive)]
                except:
                    pass

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns) > 0:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                data_frame.drop(['DepartmentID','DivisionID','IsActive','LeaveType'], inplace=True, axis=1)
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)

    def get_task_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        ids = None
        departments = None
        divisions = None
        isActive = None
        types = None
        projects = None
        phases = None
        status = None
        col_names =[
                'Employee ID', 'Employee Name', 'Project', 'Ref ID',
                'Description', 'Incident', 'Date Created',
                'Date Started', 'Target Date','Completed Date','Phase',
                'Status','Comments','Estimated Effort','Actual Effort',
                'Department','Division'
            ]
        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    ids = get_param_value(PARAM_IDS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    isActive = get_param_value(PARAM_FLAG, json_parameters)
                    types = get_param_value(PARAM_TYPES, json_parameters)
                    phases = get_param_value(PARAM_PHASES, json_parameters)
                    projects = get_param_value(PARAM_PROJECT, json_parameters)
                    status = get_param_value(PARAM_STATUS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            results = self._db_helper.get_task_view(status)
            
            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['EmployeeID'].isin(ids)]

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if types is not None and len(types)>0:
                  data_frame = data_frame[data_frame['IncidentTypeID'].isin(types)]

            if isActive is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(isActive)]
                except:
                    pass

            if projects is not None and len(projects)>0:
                  data_frame = data_frame[data_frame['ProjectID'].isin(projects)]

            if phases is not None and len(phases)>0:
                  data_frame = data_frame[data_frame['PhaseID'].isin(phases)]

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None and len(columns) > 0:        
                data_frame = data_frame[columns].head(limit)
            else:
                #data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                data_frame.drop([
                    'ProjectID','DepartmentID','DivisionID',
                    'IsActive','IncidentTypeID', 'PhaseID',
                    'TaskStatusID'
                    ], 
                    inplace=True, 
                    axis=1
                )
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)

    def get_action_list_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        ids = None
        departments = None
        divisions = None
        isActive = None
        status= None
        fys = None
        months = None
        year = None
        col_names =[
                'Action ID', 'Action', 'EmployeeID', 'Employee', 'Date Created',
                'Due Date', 'Date Closed', 'DivisionID','Division','DepartmentID',
                'Department','Status','IsActive','Month','Year','FiscalYear'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    ids = get_param_value(PARAM_IDS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    isActive = get_param_value(PARAM_FLAG, json_parameters)
                    months = get_param_value(PARAM_MONTHS, json_parameters, months)
                    year = get_param_value(PARAM_YEAR, json_parameters, year)
                    status = get_param_value(PARAM_STATUS, json_parameters)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if months is not None and arguments.auto_mode:
                months=[datetime.month]
                fys= None
                year = datetime.year
            elif arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = None
                year = None

            if fys is not None and len(fys)>0:
                results = self._db_helper.get_action_list_view(fys)
            else:
                results = self._db_helper.get_action_list_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['EmployeeID'].isin(ids)]

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if year is not None:
                data_frame = data_frame[data_frame['Year'] == year]
            
            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['Month'].isin(months)]

            if status is not None and len(status)>0:
                data_frame = data_frame[data_frame['Status'].isin(status)]

            if isActive is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(isActive)]
                except:
                    pass

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                if self._get_parameters(arguments, PARAM_COLUMNS) is None:
                    data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if columns is None:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'IsActive',
                            'EmployeeID',
                            'IsActive',
                            'Month',
                            'Year',
                            'Status',
                            'FiscalYear'

                        ], 
                        inplace=True, 
                        axis=1
                    )
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_lesson_learnt_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        ids = None
        departments = None
        divisions = None
        isActive = None
        fys = None
        months = None
        year = None
        col_names =[
                'Learning ID', 'EmployeeID', 'Employee', 'Problem',
                'Resolution','Date Created', 'Related Action',
                'DepartmentID','Department',
                'DivisionID','Division',
                'Month','Year','FiscalYear','IsActive'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    ids = get_param_value(PARAM_IDS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    isActive = get_param_value(PARAM_FLAG, json_parameters)
                    months = get_param_value(PARAM_MONTHS, json_parameters, months)
                    year = get_param_value(PARAM_YEAR, json_parameters, year)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if months is not None and arguments.auto_mode:
                months=[datetime.month]
                fys= None
                year = datetime.now().year
            elif arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = None
                year = None

            if fys is not None and len(fys)>0:
                results = self._db_helper.get_lesson_learnt_view(fys)
            else:
                results = self._db_helper.get_lesson_learnt_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['EmployeeID'].isin(ids)]

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if year is not None:
                data_frame = data_frame[data_frame['Year'] == year]
            
            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['Month'].isin(months)]

            if isActive is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(isActive)]
                except:
                    pass

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                if self._get_parameters(arguments, PARAM_COLUMNS) is None:
                    data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if columns is None:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'IsActive',
                            'EmployeeID',
                            'Month',
                            'Year',
                            'FiscalYear'
                        ], 
                        inplace=True, 
                        axis=1
                    )
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_project_billability_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        projects = None
        departments = None
        divisions = None
        months = None
        fys = None
        flag = None
        col_names =[
            'ProjectID', 'Project', 'WeekID', 'Hours',
            'IsBillable','Fiscal Year', 'DivisionID',
            'DepartmentID','Week Range','Month'
        ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    projects = get_param_value(PARAM_PROJECT, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    months= get_param_value(PARAM_MONTHS, json_parameters)
                    flag = get_param_value(PARAM_FLAG, json_parameters)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = [datetime.now().month]
                

            if fys is not None and len(fys)>0:
                results = self._db_helper.get_project_billability_view(fys)
            else:
                results = self._db_helper.get_project_billability_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if projects is not None and len(projects)>0:
                data_frame = data_frame[data_frame['ProjectID'].isin(ids)]

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['Month'].isin(months)]

            if flag is not None:
                try:
                    data_frame = data_frame[data_frame['IsBillable'] == int(flag)]
                except:
                    pass

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if arguments.action == CMD_ACTIONS[1]:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'ProjectID',
                            'IsBillable',
                            'WeekID',
                            'Month'
                        ], 
                        inplace=True, 
                        axis=1
                    )
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_employee_billability_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        projects = None
        ids = None
        departments = None
        divisions = None
        months = None
        fys = None
        flag = None
        col_names =[
                'EmployeeID', 'Employee Name', 'ProjectID', 'Project',
                'WeekID', 'Hours', 'IsBillable','Fiscal Year', 'DivisionID',
                'DepartmentID','Month', 'Week Range'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    projects = get_param_value(PARAM_PROJECT, json_parameters)
                    ids = get_param_value(PARAM_IDS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    months= get_param_value(PARAM_MONTHS, json_parameters)
                    flag = get_param_value(PARAM_FLAG, json_parameters)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = [datetime.now().month]
                
            if fys is not None and len(fys)>0:
                results = self._db_helper.get_employee_billability_view(fys)
            else:
                results = self._db_helper.get_employee_billability_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if projects is not None and len(projects)>0:
                data_frame = data_frame[data_frame['ProjectID'].isin(projects)]

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['Month'].isin(months)]

            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['EmployeeID'].isin(ids)]

            if flag is not None:
                try:
                    data_frame = data_frame[data_frame['IsBillable'] == int(flag)]
                except:
                    pass

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if arguments.action == CMD_ACTIONS[1]:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'ProjectID',
                            'IsBillable',
                            'WeekID',
                            'Month'
                        ], 
                        inplace=True, 
                        axis=1
                    )
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_concern_list_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        ids = None
        departments = None
        divisions = None
        isActive = None
        fys = None
        months = None
        flag = None
        col_names =[
                'Concern ID', 'Concern', 'Cause', 'Countermeasure',
                'RaisedByID','Raised By', 'Date Raised',
                'Date Due','DivisionID','Division',
                'DepartmentID','Department','StatusID','Status',
                'IsActive','Month','Year','Fiscal Year'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    ids = get_param_value(PARAM_IDS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    months = get_param_value(PARAM_MONTHS, json_parameters, months)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                    flag = get_param_value(PARAM_FLAG,json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if months is not None and arguments.auto_mode:
                months=[datetime.month]
                fys= None
            elif arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = None
                
            if fys is not None and len(fys)>0:
                results = self._db_helper.get_concern_list_view(fys)
            else:
                results = self._db_helper.get_concern_list_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if ids is not None and len(ids)>0:
                data_frame = data_frame[data_frame['RaisedByID'].isin(ids)]

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['Month'].isin(months)]

            if flag is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(flag)]
                except:
                    pass
            
            if arguments.auto_mode:
                data_frame.sort_values(by=['DateRaised'], inplace= True)
      
            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)

            
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                if self._get_parameters(arguments, PARAM_COLUMNS) is None:
                    data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if columns is None:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'IsActive',
                            'RaisedByID',
                            'Month',
                            'Year',
                            'Division',
                            'Department',
                            'StatusID'
                        ], 
                        inplace=True, 
                        axis=1
                    )
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_success_register_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        departments = None
        divisions = None
        isActive = None
        fys = None
        months = None
        flag = None
        col_names =[
                'SuccessID', 'RaisedByID', 'Submitted By', 'Date Submitted',
                'Participants','Details', 'Additional Information',
                'DivisionID','Division','DepartmentID','Department',
                'IsActive','Month','Year','FiscalYear'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    flag = get_param_value(PARAM_FLAG, json_parameters)
                    months = get_param_value(PARAM_MONTHS, json_parameters, months)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if months is not None and arguments.auto_mode:
                months=[datetime.month]
                fys= None
            elif arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = None
          
            if fys is not None and len(fys)>0:
                results = self._db_helper.get_success_register_view(fys)
            else:
                results = self._db_helper.get_success_register_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['Month'].isin(months)]

            if flag is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(flag)]
                except:
                    pass
                
            if arguments.auto_mode:
                data_frame.sort_values(by=['DateSubmitted'], inplace= True)

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                if self._get_parameters(arguments, PARAM_COLUMNS) is None:
                    data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if columns is None:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'IsActive',
                            'RaisedByID',
                            'Month',
                            'Year',
                            'Division',
                            'Department',
                            'SuccessID'
                        ], 
                        inplace=True, 
                        axis=1
                    )
            except Exception as ex:
                print(ex)

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_comcell_schedule_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        departments = None
        divisions = None
        fys = None
        isActive = None
        col_names =[
                'ComCellID', 'CreatedByID', 'Month','Facilitator', 
                'Minute Taker','DivisionID','Division',
                'DepartmentID','Department',
                'IsActive', 'MonthID','Fiscal Year'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    ids = get_param_value(PARAM_IDS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    isActive = get_param_value(PARAM_FLAG, json_parameters)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
             
            if fys is not None and len(fys)>0:
                results = self._db_helper.get_comcell_schedule_view(fys)
            else:
                results = self._db_helper.get_comcell_schedule_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if isActive is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(isActive)]
                except:
                    pass

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
            else:
                data_frame.sort_values(by=['MonthID'], inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                if self._get_parameters(arguments, PARAM_COLUMNS) is None:
                    data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if columns is None:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'IsActive',
                            'CreatedByID',
                            'Division',
                            'Department',
                            'MonthID',
                            'ComCellID'
                        ], 
                        inplace=True, 
                        axis=1
                    )
            except:
                pass

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_kpi_summary_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        departments = None
        divisions = None
        months = None
        fys = None
        col_names =[
                'KPIRefID', 'Subject', 'Description', 'Target',
                'Actual','Overall','MonthID','Month','Fiscal Year',
                'DivisionID','DepartmentID','DatePosted'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    months = get_param_value(PARAM_MONTHS, json_parameters, months)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = None
          
            if fys is not None and len(fys)>0:
                results = self._db_helper.get_kpi_summary_view(fys)
            else:
                results = self._db_helper.get_kpi_summary_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['MonthID'].isin(months)]

           
            data_frame.sort_values(by=['MonthID'], inplace= True)

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                if self._get_parameters(arguments, PARAM_COLUMNS) is None:
                    data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if arguments.action == CMD_ACTIONS[1]:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'MonthID',
                            'DatePosted'
                        ], 
                        inplace=True, 
                        axis=1
                    )
                else:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'DatePosted'
                        ], 
                        inplace=True, 
                        axis=1
                    )
            except Exception as ex:
                print(ex)

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def get_attendance_view_data_frame(self, arguments: CliArgument):

        columns = None
        sort_keys = None
        departments = None
        divisions = None
        months = None
        fys = None
        status = None
        isActive = None
        col_names =[
                'EmployeeID', 'Employee Name', 'Time In', 'Status',
                'StatusID','Fiscal Year',
                'DivisionID','Division','DepartmentID','Department',
                'IsActive','MonthID', 'Month'
            ] 

        try:
            if arguments.parameters is not None:
                try:
                    json_parameters = json.loads(arguments.parameters)
                    sort_keys = get_param_value(PARAM_SORT, json_parameters)
                    columns = get_param_value(PARAM_COLUMNS, json_parameters)
                    departments = get_param_value(PARAM_DEPARTMENTS, json_parameters)
                    divisions = get_param_value(PARAM_DIVISIONS, json_parameters)
                    months = get_param_value(PARAM_MONTHS, json_parameters, months)
                    fys = get_param_value(PARAM_FYS, json_parameters)
                    status = get_param_value(PARAM_STATUS, json_parameters)
                    isActive = get_param_value(PARAM_FLAG, json_parameters)
                except Exception as ex:
                    self._log(f'error reading parameters . {str(ex)}',2)
            
            if arguments.auto_mode:
                fys=[self._report_defaults['fiscal_year']]
                months = None
          
            if status is not None and len(status)>0:
                results = self._db_helper.get_attendance_view(status)
            else:
                results = self._db_helper.get_attendance_view()

            data_frame = pd.DataFrame(results.fetchall()) 
            data_frame.columns = results.keys()

            if departments is not None and len(departments)>0:
                 data_frame = data_frame[data_frame['DepartmentID'].isin(departments)]
            
            if divisions is not None and len(divisions)>0:
                  data_frame = data_frame[data_frame['DivisionID'].isin(divisions)]

            if fys is not None and len(fys)>0:
                  data_frame = data_frame[data_frame['FiscalYear'].isin(fys)]

            if months is not None and len(months)>0:
                data_frame = data_frame[data_frame['MonthID'].isin(months)]

            if isActive is not None:
                try:
                    data_frame = data_frame[data_frame['IsActive'] == int(isActive)]
                except:
                    pass

            data_frame.sort_values(by=['MonthID'], inplace= True)

            if sort_keys is not None and len(sort_keys) > 0:
                data_frame.sort_values(by=sort_keys, inplace= True)
     
            limit = int(arguments.result_limit)

            if columns is not None:        
                data_frame = data_frame[columns].head(limit)
            else:
                if self._get_parameters(arguments, PARAM_COLUMNS) is None:
                    data_frame.columns = col_names
                data_frame = data_frame.head(limit)
            
            try:
                if arguments.action == CMD_ACTIONS[1]:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID',
                            'MonthID',
                            'IsActive',
                            'StatusID'
                        ], 
                        inplace=True, 
                        axis=1
                    )
                else:
                    data_frame.drop(
                        [
                            'DepartmentID',
                            'DivisionID'
                        ], 
                        inplace=True, 
                        axis=1
                    )
            except Exception as ex:
                print(ex)

            return data_frame

        except Exception as ex:
            self._log(f'error getting data frame. {str(ex)}',2)
    
    def _retrieve_status_report_view(self, arguments: CliArgument)->None:
        
        try:
            data_frame = self.get_status_report_data_frame(arguments)

            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
       
    def _retrieve_asset_inventory_view(self, arguments: CliArgument)->None:
        
        try:
            data_frame = self.get_asset_inventory_data_frame(arguments)

            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
        
    def _retrieve_commendation_view(self, arguments: CliArgument)->None:
        
        try:
            data_frame = self.get_commendation_view_data_frame(arguments)
            col_widths = [[0,25],[1,20],[2,20],[3,20],[4,60]]
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Commendations',
                column_widths = col_widths
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
     
    def _retrieve_contact_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_contact_view_data_frame(arguments)
            col_widths = [[0,15],[1,30],[2,8],[3,15],[4,15],[5,15],[6,25],[7,30],[8,15],[9,15]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Contacts',
                column_widths=col_widths
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)

    def _retrieve_leave_summary_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_leave_summary_view_data_frame(arguments)
            col_widths = [[0,15],[1,30],[2,15],[3,12],[4,12],[5,12],[6,12],[7,10]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Leave Summary',
                column_widths=col_widths
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)

    def _retrieve_task_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_task_view_data_frame(arguments)
            col_widths = [[0,15],[1,30],[2,20],[3,12],[4,35],[5,15],[6,15],[7,15],[8,15],[9,15],[10,15],[11,15],[12,20],[13,15],[14,15],[15,15],[16,15]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Tasks',
                column_widths=col_widths,
                title='Outstanding Tasks'
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)

    def _retrieve_lesson_learnt_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_lesson_learnt_view_data_frame(arguments)
            col_widths = [[0,15],[1,30],[2,35],[3,35],[4,15],[5,35],[6,15],[7,15]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Lesson Learnt',
                column_widths=col_widths,
                title='Lesson Learnt'
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
   
    def _retrieve_action_list_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_action_list_view_data_frame(arguments)
            
            col_widths= None
            col_widths = [[0,17],[1,40],[2,25],[3,12],[4,12],[5,12],[6,15],[7,15]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Action List',
                column_widths=col_widths,
                title='Action List'
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)

    def _retrieve_project_billability_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_project_billability_view_data_frame(arguments)
            
            col_widths= None
            col_widths = [[0,17],[1,10],[2,15],[3,25]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Project Billability',
                column_widths=col_widths
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)

    def _retrieve_employee_billability_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_employee_billability_view_data_frame(arguments)
            
            col_widths= None
            col_widths = [[0,15],[1,30],[2,15],[3,10],[4,15],[5,15]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Employee Billability',
                column_widths=col_widths
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)

    def _retrieve_concern_list_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_concern_list_view_data_frame(arguments)
            col_widths = [[0,15],[1,35],[2,35],[3,35],[4,25],[5,15],[6,15],[7,15],[8,10]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = "3 C's",
                column_widths=col_widths,
                title = 'Concern, Cause and Countermeasure'
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)

    def _retrieve_success_register_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_success_register_view_data_frame(arguments)
            col_widths = [[0,25],[1,13],[2,30],[3,40],[4,40],[5,10]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Success Registers',
                column_widths=col_widths,
                title='Success Registers'
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
   
    def _retrieve_comcell_schedule_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_comcell_schedule_view_data_frame(arguments)
            col_widths = [[0,15],[1,30],[2,30],[3,15]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Schedule',
                column_widths=col_widths,
                title='Comm Cell Schedule'
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
   
    def _retrieve_kpi_summary_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_kpi_summary_view_data_frame(arguments)
            col_widths = [[0,15],[1,20],[2,30],[3,15],[4,15],[5,15],[6,16],[7,16]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'KPI',
                column_widths=col_widths,
                title='KPI Target Summary'
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)

    def _retrieve_attendance_view(self, arguments: CliArgument) ->None:
        try:
            data_frame = self.get_attendance_view_data_frame(arguments)
            col_widths = [[0,12],[1,30],[2,18],[3,10],[4,10],[5,12],[6,12],[7,12]]
            
            self._send_to_output(
                data_frame, 
                arguments.display_format, 
                arguments.output_file,
                arguments.view,
                sheet_name = 'Attendance',
                column_widths=col_widths,
                title='Attendance Report'
            )

            self._log(f"retrieved [ {len(data_frame)} ] records")
        except Exception as ex:
            self._log(f'error parsing parameter. {str(ex)}',2)
   
    def _send_to_output(
        self, 
        data_frame: pd.DataFrame, 
        format, 
        out_file, 
        view, 
        column_widths='', 
        column_names = '',
        sheet_name='Sheet1',
        title='') -> None:
        
        if out_file is None:
            file = DEFAULT_FILENAME
            out_file = file
        
        out_file = append_date_to_file_name(out_file)
        try:
            if format == OUTPUT_FORMAT[1]:
                json_file = f"{file}.json" if out_file is None else out_file
                data_frame.to_json(json_file, orient='records')
                #print(f"the file was saved : {json_file}")
            elif format == OUTPUT_FORMAT[2]:
                csv_file = f"{file}.csv" if out_file is None else out_file
                data_frame.to_csv(csv_file)
                #print(f"the file was saved : {csv_file}")
            elif format == OUTPUT_FORMAT[3]:
                excel_file = f"{file}.xlsx" if out_file is None else out_file

                if view == SUPPORTED_VIEWS[0]: # status-report
                    self._generate_raw_status_report_excel(
                        data_frame, 
                        excel_file
                    )
                elif view == SUPPORTED_VIEWS[1]: #asset-inventory
                    self._generate_raw_asset_inventory_excel(
                        data_frame, 
                        excel_file
                    )
                elif view in [
                    SUPPORTED_VIEWS[2],
                    SUPPORTED_VIEWS[3],
                    SUPPORTED_VIEWS[4],
                    SUPPORTED_VIEWS[5],
                    SUPPORTED_VIEWS[6],
                    SUPPORTED_VIEWS[7],
                    SUPPORTED_VIEWS[8],
                    SUPPORTED_VIEWS[9],
                    SUPPORTED_VIEWS[10],
                    SUPPORTED_VIEWS[11],
                    SUPPORTED_VIEWS[12],
                    SUPPORTED_VIEWS[13],
                    SUPPORTED_VIEWS[14]
                ]: 
                    self._generate_raw_excel(
                        data_frame, 
                        excel_file,
                        sheet_name = sheet_name,
                        column_widths = column_widths,
                        column_names = column_names,
                        title = title
                    )
                else:
                    self._log(f'error = {NOT_SUPPORTED}', 1)
                
            elif format == OUTPUT_FORMAT[0]:    
                print('\n') 
                print(tabulate(
                    data_frame, 
                    showindex=False, 
                    headers=data_frame.columns
                ))
                print('\n') 
            else:
                print(NOT_SUPPORTED)

        except Exception as ex:
            self._log(str(ex),2)

    def _generate_raw_status_report_excel(self, data_frame : pd.DataFrame, file_name)-> None:
        try:
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)

            drop_columns=['WeekRangeStart', 'WeekRangeId', 'ProjectId']
            column_headers = [
                'Project',
                'Project Code',
                'Rework',
                'Ref ID',
                'Description',
                'Severity',
                'Incident Type',
                'Assigned Employee',
                'Phase',
                'Status',
                'Start Date',
                'Target Date',
                'Completion Date',
                'Estimate',
                'Week Effort',
                'Actual Effort',
                'Comments',
                'Inbound Contacts',
                'Week End Date',
                'DepartmentID',
                'DivsionID'

            ]

            data_frame.drop(drop_columns, inplace=True, axis=1)
            data_frame.columns = column_headers
            total_row = len(data_frame)

            grouped_per_project = data_frame.groupby('Project')

            for key, values in grouped_per_project:
                df_per_group = pd.DataFrame(grouped_per_project.get_group(key))
                #df_per_group.columns = column_headers
                grouped_per_project_by_week = df_per_group.groupby([
                    'Week End Date'
                ])
               
                df_per_group.to_excel(
                    self._writer, 
                    sheet_name=key, 
                    index= False 
                )

                worksheet = self._writer.sheets[key]
                worksheet.set_column(0,3,12)
                worksheet.set_column(4,4,45)
                worksheet.set_column(5,5,12.5)
                worksheet.set_column(6,6,20)
                worksheet.set_column(7,7,20)
                worksheet.set_column(8,12,18)
                worksheet.set_column(13,15,12)
                worksheet.set_column(16,19,25)
                
                for col_num, value in enumerate(df_per_group.columns.values):
                    worksheet.write(0, col_num, value, main_header_format)       

            worksheet.write(
                total_row + 4, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                footer_format
            )
                 
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            #print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _generate_raw_asset_inventory_excel(self, data_frame : pd.DataFrame, file_name)-> None:
        try:
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)

            drop_columns=[
                'Comments', 
                'OtherInfo', 
                'DateAssigned',
                'DatePurchased',
                'DepartmentID',
                'DivisionID'
            ]

            column_headers = [
                'Employee ID',
                'Employee Name',
                'Department',
                'Description',
                'Manufacturer',
                'Model',
                'Serial Number',
                'Asset Tag'
            ]
            
            data_frame.drop(drop_columns, inplace=True, axis=1)
            data_frame.columns = column_headers
        
            data_frame.to_excel(
                self._writer, 
                sheet_name='Asset Inventory', 
                index= False 
            )

            worksheet = self._writer.sheets['Asset Inventory']
            worksheet.set_column(0,0,12)
            worksheet.set_column(1,1,35)
            worksheet.set_column(2,3,14)
            worksheet.set_column(4,7,22)
            
            for col_num, value in enumerate(data_frame.columns.values):
                worksheet.write(0, col_num, value, main_header_format)       

            total_row = len(data_frame)
            worksheet.write(
                total_row + 4, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                footer_format
            )
               
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            #print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _generate_raw_asset_inventory_excel(self, data_frame : pd.DataFrame, file_name)-> None:
        try:
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            
        
            data_frame.to_excel(
                self._writer, 
                sheet_name='Asset Inventory', 
                index= False 
            )

            worksheet = self._writer.sheets['Asset Inventory']
            worksheet.set_column(0,0,12)
            worksheet.set_column(1,1,35)
            worksheet.set_column(2,3,14)
            worksheet.set_column(4,7,22)
            
            for col_num, value in enumerate(data_frame.columns.values):
                worksheet.write(0, col_num, value, main_header_format)       

            total_row = len(data_frame)
            worksheet.write(
                total_row + 4, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                footer_format
            )
               
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            #print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _generate_raw_excel(
        self, 
        data_frame : pd.DataFrame, 
        file_name, 
        sheet_name = 'Sheet1', 
        column_widths='',
        column_names = '',
        title = '')-> None:

        try:
            self._writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            
            self._workbook = self._writer.book
            main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            cell_wrap_noborder = self._workbook.add_format(SHEET_CELL_WRAP_NOBORDER)
            report_title = self._workbook.add_format(SHEET_TITLE)
        
            if len(column_names) > 0:
                data_frame.columns = column_names

            data_frame.to_excel(
                self._writer, 
                sheet_name=sheet_name, 
                index= False ,
                startrow = 1 if len(title)>0 else 0
            )

            worksheet = self._writer.sheets[sheet_name]
            current_row = 0

            if len(title) >0:
                title_range = get_cell_range_address(
                    get_cell_address(0,1),
                    get_cell_address(len(column_widths) - 1,1)
                )
                worksheet.merge_range(title_range,'','')
                worksheet.write(current_row, 0, title,report_title) 
                current_row += 1 

            if column_widths != '' and len(column_widths) >0:
                for col_width in column_widths:
                    worksheet.set_column(col_width[0],col_width[0], col_width[1], cell_wrap_noborder)

            
            for col_num, value in enumerate(data_frame.columns.values):
                worksheet.write(current_row, col_num, value, main_header_format)       

            total_row = len(data_frame)
            worksheet.write(
                total_row + 4, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                footer_format
            )
               
            self._writer.save()
            # data_frame.to_excel(file_name, index=False)
            #print(f'the file was saved : {file_name}') 
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)
 
    def _load_report_settings(self) -> None:
        self._report_settings = self._config.get_section('reports')
        self._report_schedules = self._report_settings['schedules']
        self._report_defaults = self._report_settings['defaults']

    def _get_report_schedule(self, month: int) -> List[str]:
        return self._report_schedules[month-1]

    def _get_parameters(self, arguments: CliArgument, parameter_name: str):
        param_value = None
        try:
            if arguments.parameters is not None:
                json_parameters = json.loads(arguments.parameters) 
                param_value = get_param_value(PARAM_SORT, json_parameters)

            return param_value
        except:
            return None