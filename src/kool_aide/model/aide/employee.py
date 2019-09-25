import pandas as pd


class Employee:
    def __init__(self, result_set=None):
        self.id = 0
        self.custom_id = 0
        self.last_name = ''
        self.first_name = ''
        self.middle_name = ''
        self.birth_date = None
        self.position_id = 0
        self.hire_date = None
        self.status = ''
        self.image_path = ''
        self.group_id = 0
        self.department_id = 0
        self.is_active = 0
        self.division_id = 0
        self.shift_status = ''
        self.is_approved = 0

        if result_set is not None:
            self.load(result_set)

    def populate_from_data_row(self, data_row) -> None:
        pass

    def populate_from_json(self, json_data) -> None:
        self.id = json_data['EMP_ID']
        self.custom_id = json_data['WS_EMP_ID']
        self.last_name = json_data['LAST_NAME']
        self.first_name = json_data['FIRST_NAME']
        self.middle_name = json_data['MIDDLE_NAME']
        self.birth_date = json_data['BIRTHDATE']
        self.position_id = json_data['POS_ID']
        self.hire_date = json_data['DATE_HIRED']
        self.status = json_data['STATUS']
        self.image_path = json_data['IMAGE_PATH']
        self.group_id = json_data['GRP_ID']
        self.department_id = json_data['DEPT_ID']
        self.is_active = json_data['ACTIVE']
        self.division_id = json_data['DIV_ID']
        self.shift_status = json_data['SHIFT_STATUS']
        self.is_approved = json_data['APPROVED']

    def is_ok_to_add(self) -> bool:
        if self.id is not None and \
            self.custom_id is not None and \
            self.last_name is not None and \
            self.first_name is not None and \
            self.status is not None and \
            self.last_name != '' and \
            self.first_name != '' and \
            self.is_approved is not None :

                return True
        
        return False

    def load(self, result_set) -> None:
        pass