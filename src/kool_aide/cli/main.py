import argparse
import json
# from pyfiglet import Figlet

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import print_to_screen
from kool_aide.model.cli_argument import CliArgument
from kool_aide.db_access.connection import Connection
from kool_aide.cli.command_processor import CommandProcessor
from kool_aide.assets.resources.version import *
from kool_aide.assets.resources.messages import *


def log(message, level = 3):
    logger.log(f"{message} [cli.main]", level)


if __name__ == "__main__":
    config = AppSetting()
    config.load()
    logger = CustomLogger(config)
    arguments = CliArgument()

    log(f'*******************************************************************')
    log(f'starting {APP_TITLE} main cli module v{APP_VERSION} [{APP_RELEASE}]')

    
    db_connection = Connection(config, logger)
    processor = CommandProcessor(logger, config, db_connection)
    
    # region create arg parser
    parser = argparse.ArgumentParser(description=APP_DESCRIPTION, epilog=APP_EPILOG)
    parser.add_argument('action', 
                        help='the action to perform', 
                        action='store', 
                        choices=CMD_ACTIONS)
    parser.add_argument('-r', '--report',
                        help='report name to generate', 
                        action='store', 
                        dest='report_to_generate',
                        choices=REPORT_TYPES)
    parser.add_argument('-m','--model',   
                        help='the data model to use', 
                        action='store', 
                        dest='model',
                        choices=SUPPORTED_MODELS)
    parser.add_argument('-vw','--view',   
                        help='the data view to use', 
                        action='store', 
                        dest='view',
                        choices=SUPPORTED_VIEWS)
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
    parser.add_argument('--format',
                        help='result format', 
                        action='store', 
                        dest='display_format',
                        choices= OUTPUT_FORMAT)
    parser.add_argument('--limit',
                        help='limit the number of records', 
                        action='store', 
                        dest='result_limit',
                        nargs='?',
                        const=0)
    parser.add_argument('--params',
                        help='command parameters in json', 
                        action='store', 
                        dest='parameters',
                        type=str)   
    parser.add_argument('--autorun',
                        help='command is run via another application or script',
                        action='store_true',
                        dest='auto_mode')    
    parser.add_argument('--template',
                        help='generate a template',
                        action='store_true',
                        dest='template_only')   
    parser.add_argument('-v',
                        '--version',
                        action='version',
                        version=f'{APP_TITLE} [{APP_RELEASE}] v{APP_VERSION}')          
    # endregion

    # custom_fig = Figlet(font='doom')
    # print(custom_fig.renderText(APP_TITLE), end='')
   
    arguments.load_arguments(parser.parse_args())
    log(str(arguments), 4)

    print_to_screen(f'\n*** kool-aide v{APP_VERSION} ***\n', arguments.quiet_mode, False)

    # delegate the command 
    result, message = processor.delegate(arguments)
    log(f"result = {result} | {message}")

    log("end of main")