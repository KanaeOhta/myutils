from datetime import datetime
import json
import os

import openpyxl


JOINT = '.'


class ReadJson:

    def __init__(self, file):
        self.file = file


    def __iter__(self):
        with open(self.file, 'r', encoding='utf-8') as f:
            for dic in json.load(f):
                yield dic


class Convert:

    def get_file_path(self, original_path, ext):
        suffix = datetime.now().strftime('%Y%m%d%H%M%S')
        path = '{}_{}{}'.format(os.path.splitext(original_path)[0],
            suffix, ext)
        return path


    def serialize(self, dic, pref=''):
        for key, val in dic.items():
            if isinstance(val, dict):
                yield from self.serialize(val, f'{pref}{key}{JOINT}')
            elif isinstance(val, list):
                if val:
                    for i, list_val in enumerate(val):
                        if isinstance(list_val, dict):
                            yield from self.serialize(
                                list_val, f'{pref}{key}{JOINT}{str(i)}{JOINT}')
                        else:
                            yield from self.serialize(
                                {f'{pref}{key}{JOINT}{str(i)}' : list_val})
                else:
                    yield f'{pref}{key}{JOINT}{str(0)}', val
            else:
                yield f'{pref}{key}', val


class ToExcel(Convert):

    def __init__(self, json_file):
        self.excel_file = self.get_file_path(json_file, '.xlsx')
        self.json = ReadJson(os.path.abspath(json_file))
       

    def convert(self):
        records = ({k : v for k, v in self.serialize(dic)} for dic in self.json)
        self.write(records)


    def write(self, records):
        for record in records:
            print(record)

    

if __name__ == '__main__':
    # test_dic1 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3]}
    # test_dic2 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3], 'e': [{'f': 5, 'g': 6}, {'f': 100, 'g': 120}]}
    # test_dic3 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3], 'e': [{'f': 5, 'g': 6, 'h': [89, 56, 23]}, {'f': 100, 'g': 120, 'h': [70, 56, 20]}]}
    # test_dic4 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [[1, 2, 3], [4, 5, 6]]}
    # test_dic5 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [[], []]}
    # test_dic6 = {'c': {'a': 2, 'b': {'x': 5, 'y': 10}, 'e': [10, 20, 30]}}

    to_excel = ToExcel('database.json')
    print(to_excel.excel_file)
    to_excel.convert()
    # print({k : v for k, v in converter.serialize(test_dic3)})
