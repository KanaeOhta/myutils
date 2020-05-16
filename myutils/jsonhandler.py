from functools import reduce
import json


class JsonHandler:

    def __init__(self, json_file):
        self.file = json_file


    def read(self, func, json_file=None):
        json_file = self.file if json_file is None else json_file  
        with open(json_file, 'r', encoding='utf-8') as f:
            js_dicts = func(json.load(f)) 
        return js_dicts

    
    def output(self, records, output_file=None):
        output_file = self.file if output_file is None else output_file
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(
                list(records), 
                f, 
                ensure_ascii=False, 
                indent=0
            )

    
    def _flatten(self, items, key, val):
        if isinstance(val, dict):
            return {**items, **self.flatten(val)}
        elif isinstance(val, list):
            if val and isinstance(val[0], dict):
                for dic in val:
                    return {**items, **self.flatten(dic)}
            else:
                return {**items, key: ','.join(val)}
        else:
            return {**items, key: val}


    def flatten(self, d):
        return reduce(lambda new_d, kv: self._flatten(new_d, *kv), d.items(), {})


    def get_values(self, keys, dic):
        flattened = self.flatten(dic)
        return {k: v for k, v in flattened.items() if k in keys}



       