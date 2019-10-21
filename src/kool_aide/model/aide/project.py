# kool-aide/model/project.py

class Project:
    def __init__(self, result_set = None):
        self.id = 0
        self.name = ""
        self.category = 0
        self.billability = 0
        self.employee_id = 0
        self.display_flag = 0
        self.project_code = ''
        if result_set is not None:
            self.load(result_set)

    def load(self, result):
        self.id = result["PROJ_ID"]
        self.name = result["PROJ_NAME"]
        self.category = result["CATEGORY"]
        self.billability = result["BILLABILITY"]
        self.employee_id = result["EMP_ID"]
        self.display_flag = result["DSPLY_FLG"]
        self.project_code = result['PROJ_CD']

    def to_json(self):
        project = {}

        project["id"] = self.id
        project["name"] = self.name
        project["category"] = self.category
        project["billability"] = self.billability
        project["employee_id"] = self.employee_id
        project["display_flag"] = self.display_flag
        project['project_code'] = self.project_code

        return project

    def to_csv(self):
        return f"{self.id},{self.name},{self.category},{self.billability},{self.employee_id},{self.display_flag}"
    
    def populate_from_data_row(self, data_row) -> None:
        pass

    def populate_from_json(self, json_data) -> None:
        pass

    def is_ok_to_add(self) -> bool:
        if  self.category is not None and \
            self.billability is not None and \
            self.name is not None and \
            self.employee_id is not None and \
            self.display_flag is not None and \
            len(self.name) > 0 and \
            self.employee_id > 0 :

                return True
        
        return False

    def populate_from_json(self, json_data) -> None:
        self.name = json_data['PROJ_NAME']
        self.category = json_data['CATEGORY']
        self.billability = json_data['BILLABILITY']
        self.employee_id = json_data['EMP_ID']
        self.display_flag = json_data['DSPLY_FLG']
        self.project_code = json_data['PROJ_CD']
        

