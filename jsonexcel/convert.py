from functools import reduce
import json


class JsonHandler:

    def read(self):
        pass

    def _serialize(self, items, key, val, pref):
        if isinstance(val, dict):
            return {**items, **self.serialize(val, pref + key + '.')}
        elif isinstance(val, list):
            if val and isinstance(val[0], dict):
                for i, dic in enumerate(val):
                    return {**items, **self.serialize(dic, pref + key + '.' + str(i) + '.')}
            else:
                return {**items, pref + key + '.0': ','.join(val)}
        else:
            return {**items, pref + key: val}


    def serialize(self, d, pref=''):
        return reduce(lambda new_d, kv: self._serialize(new_d, *kv, pref), d.items(), {})







class ToExcel:

    def __init__(self, json_file):
        self.json_file = json_file
