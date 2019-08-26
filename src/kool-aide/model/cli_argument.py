# kool-aide/mode/cli_argument.py

class CliArgument:
    def __init__(self):
        self.action = ""
        self.model = ""
        self.input_file = ""
        self.output_file = ""
        self.user_id = ""
        self.password = ""
        self.is_inline_parameter = False
        self.is_csv_format = False  # default is json
        self.quiet_mode = False
        self.interactive_mode = False
        self.report = ""

    def load_arguments(self, result):
        self.action = result.action
        self.model = result.model
        self.input_file = result.input_file
        self.output_file = result.output_file
        self.user_id = result.user_id
        self.password = result.password
        self.is_inline_parameter = result.is_inline
        self.is_csv_format = result.is_csv  # default is json
        self.quiet_mode = result.is_csv
        self.interactive_mode = result.interactive_mode
        self.report = result.report_to_generate

    def __str__(self):
        return f"[arguments = [action : {self.action} ; model : {self.model} ; " +\
                f"input_file : {self.input_file} ; out_file : {self.output_file} ; " +\
                f"user_id : {self.user_id} ; is_inline : {self.is_inline_parameter} ; " +\
                f"is_csv : {self.is_csv_format} ; quiet_mode : {self.quiet_mode} ; " +\
                f"interactive : {self.interactive_mode} ; report_to_generate : {self.report}]]"
