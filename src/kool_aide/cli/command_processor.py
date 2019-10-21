# kool-aide/cli/command_processor.py

from kool_aide.db_access.connection import Connection

from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.app_setting import AppSetting
from kool_aide.library.constants import *
from kool_aide.library.utilities import print_to_screen

from kool_aide.model.cli_argument import CliArgument

from kool_aide.processor.report_manager import ReportManager
from kool_aide.processor.common_manager import CommonManager
from kool_aide.processor.view_manager import ViewManager
from kool_aide.processor.attendance_manager import AttendanceManager
from kool_aide.processor.employee_manager import EmployeeManager
from kool_aide.processor.department_manager import DepartmentManager
from kool_aide.processor.division_manager import DivisionManager
from kool_aide.processor.project_manager import ProjectManager


class CommandProcessor:
    def __init__(self, logger: CustomLogger, 
                    config: AppSetting, db_connection: Connection):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        
        self._log('creating component')

    def _log(self, message, level = 3):
        self._logger.log(f'{message} [cli.command_processor]', level)
 
    def delegate(self, arguments: CliArgument):
        self._log(f'delegating {str(arguments)}')

        if arguments.action == CMD_ACTIONS[0]: # create
            if arguments.model in SUPPORTED_MODELS:
                result, message = self._create(arguments)
                if result:
                    # print_to_screen('Entry/ies was created.',arguments.quiet_mode)
                    return True, f'{result} | {message}'
                else:
                    # print_to_screen('Entry/ies was not created. Check logs.',arguments.quiet_mode)
                    return False, f'{result} | {message}'
            else:
                print_to_screen('Model not supported. Check logs.',arguments.quiet_mode)
                self._log(f'model not supported : {arguments.model}')
                return False, 'model not supported'
        elif arguments.action == CMD_ACTIONS[1]: # retrieve
            if arguments.model in SUPPORTED_MODELS:
                result, message = self._retrieve_model(arguments)
                if result:
                    print_to_screen('Data retrieval done.',arguments.quiet_mode)
                    return True, f'{result} | {message}'
                else:
                    print_to_screen('Data was not retrieved. Check logs.',arguments.quiet_mode)
                    return False, f'{result} | {message}'
            elif arguments.view in SUPPORTED_VIEWS:
                result, message = self._retrieve_view(arguments)
                if result:
                    print_to_screen('Data retrieval done.',arguments.quiet_mode)
                    return True, f'{result} | {message}'
                else:
                    print_to_screen('Data was not retrieved. Check logs.',arguments.quiet_mode)
                    return False, f'{result} | {message}'
            else:
                print_to_screen('Model/view not supported. Check logs.',arguments.quiet_mode)
                self._log(f'model/view not supported : {arguments.model}')
                return False, 'model/view not supported'             
        elif arguments.action == CMD_ACTIONS[2]: # update
            return True, ''
        elif arguments.action == CMD_ACTIONS[3]: # delete
            return True, ''
        elif arguments.action == CMD_ACTIONS[4]: # gen report
            if arguments.report in REPORT_TYPES:    
                result, message = self._generate_report(arguments)

                if result:
                    print_to_screen('Report generated.',arguments.quiet_mode)
                    return True, f'{result} | {message}'
                else:
                    print_to_screen('Report was not generated. Check logs.',arguments.quiet_mode)
                    return False, f'{result} | {message}'
            else:
                print_to_screen('Report not supported. Check logs.',arguments.quiet_mode)
                self._log(f'report type not supported : {arguments.report}')
                return False, 'report type not supported'
        elif arguments.action == CMD_ACTIONS[5]: # time-in
            return self._execute(arguments)
        else:
            print_to_screen('Invalid action. Check logs.',arguments.quiet_mode)
            return False, 'invalid action'

    def _generate_report(self, arguments : CliArgument):
        self._log(f'generating report : {arguments.report}')
        self._log(f'opening connection to database ...')
        if self._connection.initialize():
            self._log(f'connected to the database ...')
            # get the mapped model for the report
            arguments.view = MAP_VIEW_TO_REPORT[arguments.report]
            report_manager = ReportManager(
                self._logger, 
                self._config, 
                self._connection, 
                arguments
            )
            return report_manager.generate(arguments)
        else:
            self._log(f'error connecting to the database', 1)
            return False, 'command execution failed'

    def _retrieve_model(self, arguments: CliArgument):
        self._log(f'retrieving model : {arguments.model}')
        self._log(f'opening connection to database ...')
        if self._connection.initialize():
            self._log(f'connected to the database ...')
            print_to_screen('Retrieving data.',arguments.quiet_mode)
            if arguments.model == SUPPORTED_MODELS[0]:
                # employee
                employee_manager = EmployeeManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return employee_manager.retrieve(arguments)
            elif arguments.model == SUPPORTED_MODELS[1]: #project
                attendance_manager = AttendanceManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return attendance_manager.retrieve(arguments)
            elif arguments.model == SUPPORTED_MODELS[3]: #project
                project_manager = ProjectManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return project_manager.retrieve(arguments)
            elif arguments.model == SUPPORTED_MODELS[4]: #departments
                department_manager = DepartmentManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return department_manager.retrieve(arguments)
            elif arguments.model == SUPPORTED_MODELS[5]: #division
                division_manager = DivisionManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return division_manager.retrieve(arguments)
            else:
                common_manager = CommonManager(
                    self._logger,
                    self._config, 
                    self._connection,
                    arguments
                )
                return common_manager.retrieve(arguments)
        else:
            self._log(f'error connecting to the database', 1)
            return False, 'command execution failed'
            
    def _retrieve_view(self, arguments: CliArgument):
        self._log(f'retrieving view : {arguments.view}')
        self._log(f'opening connection to database ...')
        
        if self._connection.initialize():
            self._log(f'connected to the database ...')
            view_manager = ViewManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
            return view_manager.retrieve(arguments)
        else:
            self._log(f'error connecting to the database', 1)
            return False, 'command execution failed'

    def _execute(self, arguments: CliArgument):
        pass

    def _create(self, arguments: CliArgument):
        self._log(f'creating data for model : {arguments.model}')
        self._log(f'opening connection to database ...')
        if self._connection.initialize():
            self._log(f'connected to the database ...')
            print_to_screen('Create process started.',arguments.quiet_mode)
            if arguments.model == SUPPORTED_MODELS[0]:
                # employee
                employee_manager = EmployeeManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return employee_manager.create(arguments)
            elif arguments.model == SUPPORTED_MODELS[3]:
                # project
                project_manager = ProjectManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return project_manager.create(arguments)
            elif arguments.model == SUPPORTED_MODELS[4]:
                # department
                department_manager = DepartmentManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return department_manager.create(arguments)
            elif arguments.model == SUPPORTED_MODELS[5]:
                # department
                division_manager = DivisionManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return division_manager.create(arguments)
            else:
                common_manager = CommonManager(
                    self._logger, 
                    self._config, 
                    self._connection, 
                    arguments
                )
                return common_manager.create(arguments)
            # elif arguments.model == SUPPORTED_MODELS[4]:
            #     # status-report
            #     status_report_manager = StatusReportManager(
            #         self._logger, 
            #         self._config, 
            #         self._connection, 
            #         arguments
            #     )
            #     return status_report_manager.retrieve(arguments)
            
            # else:
            #     common_manager = CommonManager(
            #         self._logger,
            #         self._config, 
            #         self._connection,
            #         arguments
            #     )
            #     return common_manager.retrieve(arguments)
        else:
            self._log(f'error connecting to the database', 1)
            return False, 'command execution failed'