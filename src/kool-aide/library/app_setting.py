import json

class UserSetting:
    def __init__(self, email, vsts_email="" ):
        self.email = email
        self.vsts_email=vsts_email

    def __str__(self):
        return f"user = [email : {self.email} ; vsts : {self.vsts_email} ]"


class ConnectionSetting:
    def __init__(self, server_name, uid , password, db):
        self.server_name = server_name
        self.uid = uid
        self.password = password
        self.database = db

    def __str__(self):
        return f"connection = [server : {self.server_name} ; uid = {self.uid} ; password = {self.password}]"

class CommonSetting:
    def __init__(self, log_level=3, log_location="", debug=False):
        self.log_level = log_level
        self.log_location = log_location
        self.debug_mode = debug

    def __str__(self):
        return f"common = [log_level : {self.log_level} ; log_location : {self.log_location}]"


class AppSetting:
    def __init__(self):
        self.user_setting= UserSetting("")
        self.connection_setting= ConnectionSetting("","","", "")
        self.common_setting= CommonSetting()

    def __str__(self):
        return f"app setting = [{str(self.user_setting)} ; {str(self.connection_setting)} ; {str(self.common_setting)}]"

    def load(self):
        settings = {}
        with open('kool-aide-settings.json') as json_setting:
            settings = json.load(json_setting)

        self.user_setting.email = settings['user']['email']
        self.user_setting.vsts_email = settings['user']['vsts_email']
        self.connection_setting.server_name = settings['connection_string']['server_name']
        self.connection_setting.uid = settings['connection_string']['user_id']
        self.connection_setting.password = settings['connection_string']['password']
        self.connection_setting.database = settings['connection_string']['database']
        self.common_setting.log_level = settings['common']['log_level']
        self.common_setting.log_location = settings['common']['log_location']
        self.common_setting.debug_mode = settings['common']['debug_mode'] == "True"


    def update(self):
        pass