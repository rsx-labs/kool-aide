# kool-aide/processors/report_manager.py
import pandas as pd

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import *
from kool_aide.db_access.connection import Connection
from kool_aide.model.cli_argument import CliArgument
from kool_aide.processor.view_manager import ViewManager
from kool_aide.processor.common_manager import CommonManager
from kool_aide.report.excel.status_report import StatusReport
from kool_aide.report.excel.asset_inventory_report import AssetInventoryReport
from kool_aide.report.excel.task_summary_report import TaskSummaryReport
from kool_aide.report.excel.project_billability_report import ProjectBillabilityReport
from kool_aide.report.excel.employee_billability_report import EmployeeBillabilityReport
from kool_aide.report.excel.non_billables_report import NonBillablesReport
from kool_aide.library.utilities import print_to_screen
from kool_aide.assets.resources.messages import *

class ReportManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,
                db_connection: Connection, arguments: CliArgument):

        self._logger = logger
        self._config = config
        self._connection = db_connection
        self._arguments = arguments

        self._log("creating component")

    def _log(self, message, level=3):
        self._logger.log(f"{message} [processor.report_manager]", level)

    def generate(self, arguments: CliArgument):
        self._log("report generation started")
        # get target report from argument
        if self._arguments.report == REPORT_TYPES[0]:
            # weekly status
            print_to_screen('Generating status report ...',arguments.quiet_mode)
            try:
                self._generate_status_report(arguments)
                return True, ''
            except:
                print_to_screen('Error generating report ...',arguments.quiet_mode)
                return False, REPORT_NOT_GENERATED
        elif self._arguments.report == REPORT_TYPES[1]: # asset inventory
            # asset inventory
            print_to_screen('Generating asset inventory report ...',arguments.quiet_mode)
            try:
                self._generate_asset_inventory_report(arguments)
                return True, ''
            except:
                print_to_screen('Error generating report ...',arguments.quiet_mode)
                return False, REPORT_NOT_GENERATED
        elif self._arguments.report == REPORT_TYPES[2]: # task report
            # task report
            print_to_screen('Generating task summary report ...',arguments.quiet_mode)
            try:
                self._generate_task_summary_report(arguments)
                return True, ''
            except:
                print_to_screen('Error generating report ...',arguments.quiet_mode)
                return False, REPORT_NOT_GENERATED
        elif self._arguments.report == REPORT_TYPES[3]: # project-billability
            # project-billability
            print_to_screen('Generating project billabilty report ...',arguments.quiet_mode)
            try:
                self._generate_project_billability_report(arguments)
                return True, ''
            except:
                print_to_screen('Error generating report ...',arguments.quiet_mode)
                return False, REPORT_NOT_GENERATED

        elif self._arguments.report == REPORT_TYPES[4]: # employee-billability
            # employee-billability
            print_to_screen('Generating employee billabilty report ...',arguments.quiet_mode)
            try:
                self._generate_employee_billability_report(arguments)
                return True, ''
            except:
                print_to_screen('Error generating report ...',arguments.quiet_mode)
                return False, REPORT_NOT_GENERATED

        elif self._arguments.report == REPORT_TYPES[5]: # non billable
            print_to_screen('Generating non billables report ...',arguments.quiet_mode)
            try:
                self._generate_non_billables_report(arguments)
                return True, ''
            except:
                print_to_screen('Error generating report ...',arguments.quiet_mode)
                return False, REPORT_NOT_GENERATED
        else:
            pass

    def _write_to_file(self, data):
        self._log(f"writing to {self._arguments.output_file}")

    def _generate_status_report(self, arguments: CliArgument):
        print_to_screen('Getting data ...',arguments.quiet_mode)
        view_manager = ViewManager(
            self._logger, 
            self._config, 
            self._connection
        )
        data_frame = view_manager.get_status_report_data_frame(arguments)
        
        if data_frame is not None:
            print_to_screen('Data retrieved.',arguments.quiet_mode)
            out_file = arguments.output_file
            if arguments.output_file is None:
                out_file = DEFAULT_FILENAME
            
            out_file = append_date_to_file_name(out_file)
            writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
            report = StatusReport(self._logger, self._config, data_frame, writer)

            print_to_screen(f'Writing report to : {out_file}',arguments.quiet_mode)
            return report.generate(OUTPUT_FORMAT[3])
        else:
            print_to_screen('Error generating report ...',arguments.quiet_mode)
            return False, REPORT_NOT_GENERATED
    
    def _generate_asset_inventory_report(self, arguments: CliArgument):
        print_to_screen('Getting data ...',arguments.quiet_mode)
        view_manager = ViewManager(
            self._logger, 
            self._config, 
            self._connection
        )
        data_frame = view_manager.get_asset_inventory_data_frame(arguments)
        
        if data_frame is not None:
            print_to_screen('Data retrieved.',arguments.quiet_mode)
            out_file = arguments.output_file
            if arguments.output_file is None:
                out_file = DEFAULT_FILENAME
            
            out_file = append_date_to_file_name(out_file)
            writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
            report = AssetInventoryReport(
                self._logger, 
                self._config, 
                data_frame, 
                writer
            )
            # generate excel report
            print_to_screen(f'Writing report to : {out_file}',arguments.quiet_mode)
            return report.generate(OUTPUT_FORMAT[3])
        else:
            print_to_screen('Error generating report ...',arguments.quiet_mode)
            return False, REPORT_NOT_GENERATED
    
    def _generate_task_summary_report(self, arguments: CliArgument):
        print_to_screen('Getting data ...',arguments.quiet_mode)
        view_manager = ViewManager(
            self._logger, 
            self._config, 
            self._connection
        )
        data_frame = view_manager.get_task_view_data_frame(arguments)
        
        if data_frame is not None:
            print_to_screen('Data retrieved.',arguments.quiet_mode)
            out_file = arguments.output_file
            if arguments.output_file is None:
                out_file = DEFAULT_FILENAME
            
            out_file = append_date_to_file_name(out_file)
            writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
            report = TaskSummaryReport(
                self._logger, 
                self._config, 
                data_frame, 
                writer
            )
            # generate excel report
            print_to_screen(f'Writing report to : {out_file}',arguments.quiet_mode)
            return report.generate(OUTPUT_FORMAT[3])
        else:
            print_to_screen('Error generating report ...',arguments.quiet_mode)
            return False, REPORT_NOT_GENERATED 

    def _generate_project_billability_report(self, arguments: CliArgument):
        print_to_screen('Getting data ...',arguments.quiet_mode)
        view_manager = ViewManager(
            self._logger, 
            self._config, 
            self._connection
        )
        data_frame = view_manager.get_project_billability_view_data_frame(arguments)
        
        if data_frame is not None:
            print_to_screen('Data retrieved.',arguments.quiet_mode)
            out_file = arguments.output_file
            if arguments.output_file is None:
                out_file = DEFAULT_FILENAME
            
            out_file = append_date_to_file_name(out_file)
            writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
            report = ProjectBillabilityReport(
                self._logger, 
                self._config, 
                data_frame, 
                writer
            )
            # generate excel report
            print_to_screen(f'Writing report to : {out_file}',arguments.quiet_mode)
            return report.generate(OUTPUT_FORMAT[3])
        else:
            print_to_screen('Error generating report ...',arguments.quiet_mode)
            return False, REPORT_NOT_GENERATED   

    def _generate_employee_billability_report(self, arguments: CliArgument):
        print_to_screen('Getting data ...',arguments.quiet_mode)
        view_manager = ViewManager(
            self._logger, 
            self._config, 
            self._connection
        )
        data_frame = view_manager.get_employee_billability_view_data_frame(arguments)
        
        if data_frame is not None:
            print_to_screen('Data retrieved.',arguments.quiet_mode)
            out_file = arguments.output_file
            if arguments.output_file is None:
                out_file = DEFAULT_FILENAME
            
            out_file = append_date_to_file_name(out_file)
            writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
            report = EmployeeBillabilityReport(
                self._logger, 
                self._config, 
                data_frame, 
                writer
            )
            # generate excel report
            print_to_screen(f'Writing report to : {out_file}',arguments.quiet_mode)
            return report.generate(OUTPUT_FORMAT[3])
        else:
            print_to_screen('Error generating report ...',arguments.quiet_mode)
            return False, REPORT_NOT_GENERATED 
    
    def _generate_non_billables_report(self, arguments: CliArgument):
        print_to_screen('Getting data ...',arguments.quiet_mode)
        view_manager = ViewManager(
            self._logger, 
            self._config, 
            self._connection
        )
        data_frame = view_manager.get_employee_billability_view_data_frame(arguments)
        
        if data_frame is not None:
            print_to_screen('Data retrieved.',arguments.quiet_mode)
            out_file = arguments.output_file
            if arguments.output_file is None:
                out_file = DEFAULT_FILENAME
            
            out_file = append_date_to_file_name(out_file)
            writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
            report = NonBillablesReport(
                self._logger, 
                self._config, 
                data_frame, 
                writer
            )
            # generate excel report
            print_to_screen(f'Writing report to : {out_file}',arguments.quiet_mode)
            return report.generate(OUTPUT_FORMAT[3])
        else:
            print_to_screen('Error generating report ...',arguments.quiet_mode)
            return False, REPORT_NOT_GENERATED 
    