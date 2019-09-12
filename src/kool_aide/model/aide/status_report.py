# kool-aide/model/aide/status_report.py

class StatusReport:
    def __init__(self, result_set = None):
        self.project = ''
        self.project_code = ''
        self.rework = ''
        self.reference_id = ''
        self.description = ''
        self.severity = ''
        self.incident_type = ''
        self.assigned_employee = ''
        self.phase = ''
        self.completed_date = ''
        self.start_date = ''
        self.target_date = ''
        self.effort_estimate = ''
        self.actual_effort = ''
        self.actual_week_effort = ''
        self.comments = ''
        self.inbound_contacts = ''
        self.week_range_id = ''
        self.week_start = ''
        self.week_end = ''
        self.status = ''


        if result_set is not None:
            self.load(result_set)

    def load(self, result):
        self.project = result['Project']
        self.project_code = result['ProjectCode']
        self.rework = result['Rework']
        self.reference_id = result['ReferenceID']
        self.description = result['Description']
        self.severity = result['Severity']
        self.incident_type = result['IncidentType']
        self.assigned_employee = result['AssignedEmployee']
        self.phase = result['Phase']
        self.completed_date = result['CompletedDate']
        self.start_date = result['DateStarted']
        self.target_date = result['TargetDate']
        self.effort_estimate = result['EffortEstimate']
        self.actual_effort = result['ActualEffort']
        self.actual_week_effort = result['ActualWeekEffort']
        self.comments = result['Commentst']
        self.inbound_contacts = result['InboundContacts']
        self.week_range_id = result['WeekRangeId']
        self.week_start = result['WeekRangeStart']
        self.week_end = result['WeekRangeEnd']
        self.status = result['Status']

    def to_json(self):
        data = {}

        data['project'] = self.project 
        data['project_code'] = self.project_code 
        data['rework'] = self.rework 
        data['reference_id'] = self.reference_id 
        data['description'] = self.description 
        data['severity'] = self.severity 
        data['incident_type'] = self.incident_type 
        data['assigned_employee'] = self.assigned_employee 
        data['phase'] = self.phase 
        data['completed_date'] = self.completed_date 
        data['start_date'] = self.start_date 
        data['target_date'] = self.target_date 
        data['effort_estimate'] = self.effort_estimate 
        data['actual_effort'] = self.actual_effort 
        data['actual_week_effort'] = self.actual_week_effort 
        data['comments'] = self.comments 
        data['inbound_contacts'] = self.inbound_contacts 
        data['week_range_id'] = self.week_range_id 
        data['week_start'] = self.week_start 
        data['week_end'] = self.week_end 
        data['status'] = self.status 
       
        return data

    def to_csv(self):
        pass
    