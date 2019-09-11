import argparse
import json

from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger
from ..library.constants import *
from ..model.cli_argument import CliArgument
from ..db_access.connection import Connection
from .command_processor import CommandProcessor
from ..assets.resources.version import *
from ..assets.resources.messages import *


def log(message, level = 3):
    logger.log(f"{message} [Main]", level)


if __name__ == "__main__":
    config = AppSetting()
    config.load()
    logger = CustomLogger(config)
    arguments = CliArgument()
    db_connection = Connection(config, logger)
    processor = CommandProcessor(logger, config, db_connection)
    
    log(f'starting {APP_TITLE} main cli module v{APP_VERSION} [{APP_RELEASE}]')
    
    # region create arg parser
    parser = argparse.ArgumentParser(description=APP_DESCRIPTION, epilog=APP_EPILOG)
    parser.add_argument('action', 
                        help='the action to perform', 
                        action='store', 
                        choices=CMD_ACTIONS)
    parser.add_argument('-r',
                        help='report name to generate', 
                        action='store', 
                        dest='report_to_generate',
                        choices=REPORT_TYPES)
    parser.add_argument('-m',   
                        help='the data model to use', 
                        action='store', 
                        dest='model',
                        choices=SUPPORTED_MODELS)
    parser.add_argument('--input',  
                        help='the input file', 
                        action='store', 
                        dest='input_file')
    parser.add_argument('--output',  
                        help='the output file', 
                        action='store', 
                        dest='output_file')
    parser.add_argument('--uid',  
                        help='the user id', 
                        action='store', 
                        dest='user_id')
    parser.add_argument('--password',  
                        help='user password', 
                        action='store', 
                        dest='password')
    parser.add_argument('--quiet',  
                        help='quiet mode', 
                        action='store_true', 
                        dest='quiet_mode')
    parser.add_argument('--interactive',  
                        help='interactive mode', 
                        action='store_true', 
                        dest='interactive_mode')
    parser.add_argument('-f', '--format',
                        help='result format', 
                        action='store', 
                        dest='display_format',
                        choices= DISPLAY_FORMAT)
    parser.add_argument('-l', '--limit',
                        help='limit the number of records', 
                        action='store', 
                        dest='result_limit',
                        nargs='?',
                        const=0)
    parser.add_argument('-p', '--params',
                        help='command parameters in json', 
                        action='store', 
                        dest='parameters',
                        type=str)                    
    # endregion

    result = parser.parse_args()
    log(str(result), 4)

    arguments.load_arguments(result)
    log(str(arguments), 4)

    # start delegating if db is properly initialized
    if db_connection.initialize():
        log("we are connected to the db")
        result, message = processor.delegate(arguments)
        log(f"result = {result} | {message}")

    log("end of main")