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
            self._generate_status_report(arguments)
            return True, ''
        elif self._arguments.report == REPORT_TYPES[1]: # asset inventory
            # asset inventory
            self._generate_asset_inventory_report(arguments)
            return True, ''
        else:
            pass

    def _write_to_file(self, data):
        self._log(f"writing to {self._arguments.output_file}")

    def _generate_status_report(self, arguments: CliArgument):
        view_manager = ViewManager(
            self._logger, 
            self._config, 
            self._connection
        )
        data_frame = view_manager.get_status_report_data_frame(arguments)
        
        out_file = arguments.output_file
        if arguments.output_file is None:
            out_file = DEFAULT_FILENAME
        
        out_file = append_date_to_file_name(out_file)
        writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
        report = StatusReport(self._logger, self._config, data_frame, writer)

        return report.generate(OUTPUT_FORMAT[3])
    
    def _generate_asset_inventory_report(self, arguments: CliArgument):
        view_manager = ViewManager(
            self._logger, 
            self._config, 
            self._connection
        )
        data_frame = view_manager.get_asset_inventory_data_frame(arguments)
        
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
        return report.generate(OUTPUT_FORMAT[3])

    
        



    