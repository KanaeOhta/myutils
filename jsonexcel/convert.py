from collections import defaultdict
from datetime import datetime
import json
import os
import re

# import openpyxl
from xlsxwriter.workbook import Workbook


JOINT = '.'
MAIN = 'main'


class ReadJson:

    def __init__(self, file):
        self.file = file


    def __iter__(self):
        with open(self.file, 'r', encoding='utf-8') as f:
            for dic in json.load(f):
                yield dic


class ExcelSheet:

    def __init__(self, row=0, index=''):
        self.row = row
        self.index = index


    @property
    def column(self):
        pass

    
    def __call__(self, index):
        if self.index != index:
            self.index = index
            self.row += 1
        return self.row


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
                                list_val, f'{pref}{key}{JOINT}{i}{JOINT}')
                        else:
                            yield from self.serialize(
                                {f'{pref}{key}{JOINT}{i}' : list_val})
                else:
                    yield f'{pref}{key}{JOINT}{str(0)}', val
            else:
                yield f'{pref}{key}', val

    
    def parse_json(self, dic, pref='', group=MAIN):
        for key, val in dic.items():
            if isinstance(val, dict):
                yield from self.parse_json(val, f'{pref}{key}{JOINT}', group)
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
            self.excel_format = {key.replace(group+'.', '') : group \
                for key, group in self.excel_format.items()}


    def convert(self):
        self.get_excel_format()
        records = ({k : v for k, v in self.serialize(dic)} for dic in self.json)
        self.output(records)


    def set_sheets(self, wb):
        self.sheets = {val: ExcelSheet() for val in self.excel_format.values()}    
        for sh_name in self.sheets.keys():
            keys = [key for key, val in self.excel_format.items() if val == sh_name]
            sheet = wb.get_worksheet_by_name(sh_name)
            if sheet is None:
                sheet = wb.add_worksheet(sh_name)
            for col, key in enumerate(keys, 1):
                sheet.write(0, col, key)


    def write(self, sheet, row, index, value):
        # if value.isdecimal():
        #     sheet.write_number()
        pass


    def output(self, records):
        pattern = re.compile(r'\D+$')
        with Workbook(self.excel_file) as wb:
            self.set_sheets(wb)
            for i, record in enumerate(records, 1):
                for key, val in record.items():
                    col_name = pattern.findall(key)
                    if col_name:
                        col_name = col_name[0]
                        index = f'{i}_{key[:-len(col_name)]}'
                        col_name = col_name.lstrip('.')
                    else:
                        col_name = key
                        index = i
                    row = self.sheets(sh_name)(index)
                    sh_name = self.excel_format(col_name)
                    sheet = wb.add_worksheet(sh_name)
                    self.write(sheet, row, index, val)

                    
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
