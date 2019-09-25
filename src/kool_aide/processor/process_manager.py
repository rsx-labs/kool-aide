# kool-aide/processors/process_manager.py
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.app_setting import AppSetting


class ProcessManager:
    def __init__(self, logger: CustomLogger, config: AppSetting,  
                    name ='processor.process_manager'):
        self._name = name
        self._logger = logger
        self._config = config
    
     def _log(self, message, level=3) -> None:
        self._logger.log(f"{message} [{self._name}]", level)