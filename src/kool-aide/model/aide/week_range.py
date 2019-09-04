# kool-aide/model/week_range.py

class WeekRange:
    def __init__(self, result_set = None):
        self.id = 0
        # self.week_id
        self.start = 0
        self.end = 0
    
        if result_set is not None:
            self.load(result_set)

    def load(self, result):
        self.id = result["WEEK_ID"]
        self.start = result["WEEK_START"]
        self.end = result["WEEK_END"]

    def to_json(self):
        data = {}

        data["id"] = self.id
        data["start"] = str(self.start)
        data["end"] = str(self.end)

        return data

    def to_csv(self):
        return f"{self.id},{self.start},{self.end}"
    