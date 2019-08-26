# library/custom_logger.py
from datetime import datetime
import os
import logging
from logging.handlers import RotatingFileHandler
from logging import Formatter

from .app_setting import AppSetting

class CustomLogger:
    """ This is a custom logger for logging operations """
    
    def __init__(self,config : AppSetting):
        # log levels based on python logging library
        # 0-NOTSET; 10-DEBUG ; 20 - INFO ; 30 - WARN ; 40 - ERROR ; 50 - CRITICAL
        
        self._settings = config
        # basic logging
        # logging.basicConfig(format='%(asctime)s - %(levelname)s - %(name)s - %(message)s',
        #                     datefmt='%d-%b-%y %H:%M:%S',
        #                     level= to_integer(settings.log_level))

        # if self._settings.log_location == "":
        #     self._settings.log_location = os.getcwd()

        # path = os.path.join(self._settings.log_location,"eva.log")
        
      
        path = os.path.join(self._settings.common_setting.log_location, "kool-aide.log")

        self._logger = logging.getLogger("kool-aide")
        log_formatter = Formatter(fmt='%(asctime)s [%(levelname)s] : %(message)s  [%(threadName)s]',
                                  datefmt='%d-%b-%y %H:%M:%S')

        # console
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(log_formatter)

        # create log file of max size 1MB and keep the last 10
        file_handler = RotatingFileHandler(path, maxBytes=1048576, backupCount = 10)

        file_handler.setFormatter(log_formatter)

        self._logger.addHandler(file_handler)

        # if self._settings.app_debug_mode:
        #     # if debug mode, override log level and set to 10
        #     self._settings.log_level = 10
        if self._settings.common_setting.debug_mode:
            self._logger.addHandler(console_handler)
            
       
        #     # basic logging
        #     debug_log = os.path.join(config.app_log_location, debug_name)
               
        #     logging.basicConfig(format='%(asctime)s - %(levelname)s - %(name)s - %(message)s [%(threadName)s]',
        #                         datefmt='%d-%b-%y %H:%M:%S',
        #                         level= logging.DEBUG,
        #                         filename= debug_log)

        self._logger.setLevel(int(self._settings.common_setting.log_level))
    
    def log(self, message, level=3):
        """ 
        This is the main logging function, it logs a message based on the level

        Default Log Level if not provided = 3 (info)
        """
        
        # levels : 1 - ERROR ; 2 - WARNING ; 3 - INFO ; 4 - DEBUG
        if level == 1:
            self._logger.error(f"{message}", exc_info = True)
        elif level == 2:
            self._logger.warning(f"{message}")
        elif level == 3:
            self._logger.info(message)
        else: # level = 4 and beyond
            self._logger.debug(f"{message}")

    
