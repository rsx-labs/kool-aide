# kool-aide/model/project.py

class Project:
    def __init__(self, result_set = None):
        self.id = 0
        self.name = ""
        self.category = 0
        self.billability = 0
        self.employee_id = 0
        self.display_flag = 0
        if result_set is not None:
            self.load(result_set)

    def load(self, result):
        self.id = result["PROJ_ID"]
        self.name = result["PROJ_NAME"]
        self.category = result["CATEGORY"]
        self.billability = result["BILLABILITY"]
        self.employee_id = result["EMP_ID"]
        self.display_flag = result["DSPLY_FLG"]

    def to_json(self):
        project = {}

        project["id"] = self.id
        project["name"] = self.name
        project["category"] = self.category
        project["billability"] = self.billability
        project["employee_id"] = self.employee_id
        project["display_flag"] = self.display_flag

        return project

    def to_csv(self):
        return f"{self.id},{self.name},{self.category},{self.billability},{self.employee_id},{self.display_flag}"
    