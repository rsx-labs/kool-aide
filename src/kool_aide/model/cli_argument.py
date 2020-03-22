# kool-aide/mode/cli_argument.py

class CliArgument:
    def __init__(self):
        self.action = ""
        self.model = ""
        self.input_file = ""
        self.output_file = ""
        self.user_id = ""
        self.password = ""
        self.is_csv_format = False  # default is json
        self.quiet_mode = False
        self.interactive_mode = False
        self.report = ""
        self.display_format = ''
        self.result_limit = 0
        self.parameters = {}
        self.auto_mode = False
        self.view = ''
        self.template_only = False

    def load_arguments(self, result) -> None:
        self.action = result.action
        self.model = result.model
        self.input_file = result.input_file
        self.output_file = result.output_file
        self.user_id = result.user_id
        self.password = result.password
        # self.is_csv_format = result.is_csv  # default is json
        self.quiet_mode = result.quiet_mode
        self.interactive_mode = result.interactive_mode
        self.report = result.report_to_generate
        self.display_format = 'screen' if result.display_format is None else result.display_format
        self.result_limit =10000 if result.result_limit is None else result.result_limit
        self.parameters = result.parameters
        self.auto_mode = result.auto_mode if result.auto_mode is not None else False
        self.view = result.view
        self.template_only = result.template_only

        if self.auto_mode:
            self.quiet_mode = True

    def __str__(self):
        return f"[arguments = [action : {self.action} ; model : {self.model} ; " +\
                f"input_file : {self.input_file} ; out_file : {self.output_file} ; " +\
                f"user_id : {self.user_id} ;  " +\
                f"is_csv : {self.is_csv_format} ; quiet_mode : {self.quiet_mode} ; " +\
                f"interactive : {self.interactive_mode} ; report_to_generate : {self.report} ; " +\
                f"display_format : {self.display_format} ; result_limit : {self.result_limit} ; " + \
                f"view : {self.view}: params : {self.parameters}]]"
