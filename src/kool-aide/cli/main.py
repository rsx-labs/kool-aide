import argparse

from ..library.app_setting import AppSetting
from ..library.custom_logger import CustomLogger
from ..library.constants import *
from ..model.cli_argument import CliArgument
from ..db_access.connection import Connection
from .command_processor import CommandProcessor
from ..assets.resources.version import *


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
    parser = argparse.ArgumentParser()
    parser.add_argument('-a', 
                        help='the action to perform', 
                        action='store', 
                        dest='action',
                        choices=CMD_ACTIONS,
                        required=True)
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
    parser.add_argument('--inline',  
                        help='mark the parameters as inline json', 
                        action='store_true', 
                        dest='is_inline')
    parser.add_argument('--csv',  
                        help='input file is csv. the default is json.', 
                        action='store_true', 
                        dest='is_csv')
    parser.add_argument('-u','--uid',  
                        help='the user id', 
                        action='store', 
                        dest='user_id')
    parser.add_argument('-p','--password',  
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