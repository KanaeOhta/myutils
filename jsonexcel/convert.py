from collections import namedtuple
from datetime import datetime
import json
import os
import re

# import openpyxl
from xlsxwriter.workbook import Workbook


JOINT = '.'
HYPHEN = '-'
MAIN = 'main'


class ReadJson:

    def __init__(self, file):
        self.file = file


    def __iter__(self):
        with open(self.file, 'r', encoding='utf-8') as f:
            for dic in json.load(f):
                yield dic


Cell = namedtuple('Cell', 'key idx value')


class ExcelSheet:

    def __init__(self, sheet, keys, row=0, index=''):
        self._row = row
        self._index = index
        self.sheet = sheet
        self.set_keys(keys)


    def set_keys(self, keys):
        self.keys = {}
        for col, key in enumerate(keys, 1):
            self.keys[key] = col
            self.write(0, col, key)


    def column(self, key):
        return self.keys[key]


    def row(self, index):
        if self._index != index:
            self._index = index
            self._row += 1
        return self._row


    def write(self, row, col, value, index=''):
        if index:
            self.sheet.write(row, 0, index)
        if not value:
            self.sheet.write_blank(row, col, value)
        elif type(value) == bool:
            self.sheet.write_boolean(row, col, value)
        elif isinstance(value, str):
            self.sheet.write_string(row, col, value)
        else:
            self.sheet.write_number(row, col, value)

       
class Convert:

    def get_file_path(self, original_path, ext):
        suffix = datetime.now().strftime('%Y%m%d%H%M%S')
        path = '{}_{}{}'.format(os.path.splitext(original_path)[0],
            suffix, ext)
        return path


    def serialize(self, dic, idx, pref=''):
        for key, val in dic.items():
            if isinstance(val, dict):
                yield from self.serialize(val, 
                    idx, f'{pref}{key}{JOINT}')
            elif isinstance(val, list):
                if val:
                    for i, list_val in enumerate(val):
                        if isinstance(list_val, dict):
                            yield from self.serialize(list_val, 
                                f'{idx}{HYPHEN}{i}', f'{pref}{key}{JOINT}')
                        else:
                            yield from self.serialize({f'{pref}{key}{JOINT}{i}' : list_val}, idx)
                else:
                    yield Cell(f'{pref}{key}{JOINT}{str(0)}', idx, val)
            else:
                yield Cell(f'{pref}{key}', idx, val)

    
    def parse_json(self, dic, pref='', group=MAIN):
        for key, val in dic.items():
            if isinstance(val, dict):
                yield from self.parse_json(
                    val, f'{pref}{key}{JOINT}', group)
            elif isinstance(val, list):
                if val:
                    for i, list_val in enumerate(val):
                        if isinstance(list_val, dict):
                            yield from self.parse_json(
                                list_val, f'{pref}{key}{JOINT}', f'{pref}{key}')
                        else:
                            yield from self.parse_json(
                                {f'{pref}{key}{JOINT}{str(i)}' : list_val}, group=group)
            else:
                yield group, f'{pref}{key}'


class ToExcel(Convert):

    def __init__(self, json_file):
        self.excel_file = self.get_file_path(json_file, '.xlsx')
        self.json = ReadJson(os.path.abspath(json_file))
        self.excel_format = {}
        self.sheets = None
       

    def get_excel_format(self):
        if not self.excel_format:
            for dic in self.json:
                for group, key in self.parse_json(dic): 
                    if key not in self.excel_format.keys():
                        self.excel_format[key] = group


    def convert_all(self):
        self.get_excel_format()
        records = ((cell for cell in self.serialize(dic, str(i))) \
            for i, dic in enumerate(self.json, 1))
        self.output(records)


    def set_sheets(self, wb):
        self.sheets = {}
        for _, sh_name in self.excel_format.items():
            if sh_name not in self.sheets:  
                self.sheets[sh_name] = ExcelSheet(
                    wb.add_worksheet(sh_name),
                    [key for key, val in self.excel_format.items() if val == sh_name]
                )
                            

    def write(self, sh_name, cell):
        sheet = self.sheets[sh_name]
        col = sheet.column(cell.key)
        row = sheet.row(cell.idx)
        sheet.write(row, col, cell.value, cell.idx)


    def output(self, records):
        with Workbook(self.excel_file) as wb:
            self.set_sheets(wb)
            for record in records:
                for cell in record:
                    sh_name = self.excel_format.get(cell.key)
                    if sh_name:
                        self.write(sh_name, cell)
                    

if __name__ == '__main__':
    # test_dic1 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3]}
    # test_dic2 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3], 'e': [{'f': 5, 'g': 6}, {'f': 100, 'g': 120}]}
    # test_dic3 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3], 'e': [{'f': 5, 'g': 6, 'h': [89, 56, 23]}, {'f': 100, 'g': 120, 'h': [70, 56, 20]}]}
    # test_dic4 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [[1, 2, 3], [4, 5, 6]]}
    # test_dic5 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [[], []]}
    # test_dic6 = {'c': {'a': 2, 'b': {'x': 5, 'y': 10}, 'e': [10, 20, 30]}}

    to_excel = ToExcel('database.json')
    print(to_excel.excel_file)
    to_excel.convert_all()
    # print({k : v for k, v in converter.serialize(test_dic3)})
