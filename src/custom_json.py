import json

from datetime import date

class JSONEncoder(json.JSONEncoder):
    def default(self, o):
        if type(o) == date:
            return o.isoformat()
        return super().default(o)