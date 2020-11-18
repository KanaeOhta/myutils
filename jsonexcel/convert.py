from collections import namedtuple
from datetime import datetime
import itertools
import json
import os

import openpyxl
from xlsxwriter.workbook import Workbook


DOT = '.'
HYPHEN = '-'
MAIN = 'main'


class NoMoreRecord(Exception):
    pass


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


class ReadOnlySheet:

    def __init__(self, sheet):
        self.sheet = sheet
        self.main_key = self.sheet.title.split(DOT)[0]
        self.keys = self.set_keys()
        self.max_col = len(self.keys) + 1
        self.max_row = self.sheet.max_row
        self.row = 2
  

    def is_empty(self, cell):
        return cell.value is None or not str(cell.value).strip()    
    

    def set_keys(self):
        row = self.sheet[1]
        return tuple(cell.value for cell in row if not self.is_empty(cell))
       

    def read(self, serial):
        if self.row > self.max_row:
            raise NoMoreRecord() 
        min_row = self.row
        for row in self.sheet.iter_rows(
                min_row=min_row, max_row = self.max_row, min_col=1, max_col=self.max_col):
            a_cell = row[0]
            # when formatted cell is found out of data area
            if self.is_empty(a_cell):
                raise NoMoreRecord()
            idx = a_cell.value
            if idx.split(HYPHEN)[0] != serial:
                return
            self.row += 1
            for cell, key in zip(row[1:], self.keys):
                yield Cell(key, idx, cell.value)


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
                    idx, f'{pref}{key}{DOT}')
            elif isinstance(val, list):
                if val:
                    for i, list_val in enumerate(val):
                        if isinstance(list_val, dict):
                            yield from self.serialize(list_val, 
                                f'{idx}{HYPHEN}{i}', f'{pref}{key}{DOT}')
                        else:
                            yield from self.serialize({f'{pref}{key}{HYPHEN}{i}' : list_val}, idx)
                else:
                    yield Cell(f'{pref}{key}{DOT}{str(0)}', idx, val)
            else:
                yield f'{pref}{key}', idx, val

    
    def parse_json(self, dic, pref='', group=MAIN):
        for key, val in dic.items():
            if isinstance(val, dict):
                yield from self.parse_json(
                    val, f'{pref}{key}{DOT}', group)
            elif isinstance(val, list):
                if val:
                    for i, list_val in enumerate(val):
                        if isinstance(list_val, dict):
                            yield from self.parse_json(
                                list_val, f'{pref}{key}{DOT}', f'{pref}{key}')
                        else:
                            yield from self.parse_json(
                                {f'{pref}{key}{HYPHEN}{i}' : list_val}, group=group)
            else:
                yield group, f'{pref}{key}'


    def set_container_to_list(self, li, idx, container):
        try:
            li[idx]
        except IndexError:
            li.append(container())


    def set_container_to_dict(self, dic, key, container):
        if key not in dic:
            dic[key] = container()


    def _deserialize(self, new_dic, keys, idxes, val):
        if not idxes:
            if len(keys) == 1:
                if HYPHEN in keys[0]:
                    temp_keys = keys[0].split(HYPHEN)
                    self.set_container_to_dict(new_dic, temp_keys[0], list)
                    if val:
                        li = new_dic[temp_keys[0]]
                        li.append(val)
                    return
                new_dic[keys[0]] = val
                return
            else:
                self.set_container_to_dict(new_dic, keys[0], dict)
                self._deserialize(new_dic[keys[0]], keys[1:], idxes, val)
        else:
            self.set_container_to_dict(new_dic, keys[0], list)
            li = new_dic[keys[0]]
            self.set_container_to_list(li, idxes[0], dict)
            self._deserialize(li[idxes[0]], keys[1:], idxes[1:], val)
       

    def deserialize(self, dic):
        new_dic = {}
        for (key_str, idx_str), val in dic.items():
            keys =key_str.split(DOT)
            idxes = [int (idx) for idx in idx_str.split(HYPHEN)]
            self._deserialize(new_dic, keys, idxes[1:], val)
        return new_dic


class ToExcel(Convert):

    def __init__(self, json_file):
        json_file = os.path.abspath(json_file)
        self.excel_file = self.get_file_path(json_file, '.xlsx')
        self.json = ReadJson(json_file)
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
        records = ((Cell(key, idx, val) for key, idx, val in self.serialize(dic, str(i))) \
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


class FromExcel(Convert):

    def __init__(self, excel_file):
        self.excel_file = os.path.abspath(excel_file)
        self.sheets = None
        
        
    def set_sheets(self, wb):
        self.sheets = tuple(ReadOnlySheet(sh) for sh \
            in wb if sh.cell(row=2, column=1).value)
   

    def convert(self):
        wb = openpyxl.load_workbook(self.excel_file)
        self.set_sheets(wb)
        for item in self.read():
            print(item)
        wb.close()

    
    def read(self):
        for i in itertools.count(1):
            dic = {}
            try:
                for sheet in self.sheets:
                    dic = {
                        **dic,
                        **{(cell.key, cell.idx): cell.value for cell in sheet.read(str(i))}
                    }
                yield self.deserialize(dic)
            
            except NoMoreRecord:
                break

        
if __name__ == '__main__':
    # test_dic1 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3]}
    # test_dic2 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3], 'e': [{'f': 5, 'g': 6}, {'f': 100, 'g': 120}]}
    # test_dic3 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3], 'e': [{'f': 5, 'g': 6, 'h': [89, 56, 23]}, {'f': 100, 'g': 120, 'h': [70, 56, 20]}]}
    # test_dic4 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [[1, 2, 3], [4, 5, 6]]}
    # test_dic5 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [[], []]}
    # test_dic6 = {'c': {'a': 2, 'b': {'x': 5, 'y': 10}, 'e': [10, 20, 30]}}

    # to_excel = ToExcel('database.json')
    # to_excel = ToExcel('dict10.json')
    # print(to_excel.excel_file)
    # to_excel.convert_all()
    # print({k : v for k, v in converter.serialize(test_dic3)})
    # path = r"C:\Users\kanae\OneDrive\myDevelopment\JsonExcel\database_20201114203603.xlsx"
    path = r"C:\Users\kanae\OneDrive\myDevelopment\JsonExcel\dict10_20201116220255.xlsx"
    # path = r"C:\Users\kanae\OneDrive\myDevelopment\JsonExcel\dict10_dummy.xlsx"
    from_excel = FromExcel(path)
    from_excel.convert()
