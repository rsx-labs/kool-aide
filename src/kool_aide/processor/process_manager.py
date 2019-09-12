# kool-aide/processors/process_manager.py
from kool_aide.library.custom_logger import CustomLogger


class ProcessManager:
    def __init__(self, logger: CustomLogger, name ='process_manager'):
        self.name = name
        self._logger = logger