from platformdirs import user_config_dir
import os
import json
import sys

from crypt import encode, decode
import custom_json

class Settings:

    _data = dict()

    def __init__(self, noload=False):
        if noload: 
            return
        self.load()

    @classmethod
    def get_filepath(cls) -> str:
        filename = "settings.json"
        config_dir = user_config_dir("elexam", ensure_exists=True)
        return os.path.join(config_dir, filename)

    def get_crypted(self, param):
        value = self[param]
        if not value: return None
        decoded = decode("&(gsUDH)ds3@", value)
        try:
            return json.loads(decoded)
        except:
            return decoded

    @classmethod
    def load(cls) -> dict:
        filepath = cls.get_filepath()

        if not os.path.isfile(filepath) or os.path.getsize(filepath) == 0:
            cls._data = dict()
            return cls._data

        with open(filepath, 'r') as f:
            cls._data = json.load(f)
        return cls._data

    def save(self) -> bool:
        return self.write(self.__class__._data)

    def set_crypted(self, param, value):
        if type(value) != str:
            value = json.dumps(value)
        value = encode("&(gsUDH)ds3@", value)
        self[param] = value

    @classmethod
    def update(cls, key, value) -> bool:
        settings = cls.load()
        settings[key] = value
        return cls.write(settings)

    @classmethod
    def write(cls, settings: dict) -> bool:
        with open(cls.get_filepath(), 'w') as f:
            json.dump(settings, f, indent=4, cls=custom_json.JSONEncoder)
        cls._data = settings
        return True

    def __len__(self):
        return len(self.__class__._data)

    def __iter__(self):
        return self.__class__._data.__iter__()

    def __contains__(self, item):
        return self.__class__._data.__contains__(item)

    def __getitem__(self, key):
        # self.load()
        if not key in self.__class__._data:
            return None
        return self.__class__._data.__getitem__(key)

    def __setitem__(self, key, value):
        self.__class__._data.__setitem__(key, value)
        self.write(self.__class__._data)

    def __delitem__(self, key):
        if not key in self.__class__._data: return
        self.__class__._data.__delitem__(key)
        self.write(self.__class__._data)

