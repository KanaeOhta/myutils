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
        table = str.maketrans({HYPHEN: '_', DOT: '_'})
        self._replace = lambda x: x.translate(table)
       

    def replace(self, obj, func):
        """Replace '-' and '.' to '_', because they are used when
           flatten and unflatten dict.
        """
        if isinstance(obj, dict):
            return {func(key): self.replace(val, func) for key, val in obj.items()}
        elif isinstance(obj, list):
            return [self.replace(item, func) for item in obj]
        else:
            return obj


    def __iter__(self):
        with open(self.file, 'r', encoding='utf-8') as f:
            for dic in json.load(f):
                yield self.replace(dic, self._replace)


Cell = namedtuple('Cell', 'key idx value')


class ExcelSheet:

    def __init__(self, sheet):
        self.sheet = sheet


    def set_keys(self, *args):
        raise NotImplementedError()


class WritingSheet(ExcelSheet):

    def __init__(self, sheet, keys, row=0, index=''):
        super().__init__(sheet)
        self._row = row
        self._index = index
        self.set_keys(keys)


    def set_keys(self, keys):
        self.keys = {}
        for col, key in enumerate(keys, 1):
            self.keys[key] = col
            self.write(0, col, key)


    def add_column(self, key):
        self.write(0, len(self.keys) + 1, key)


    def column(self, key):
        return self.keys.get(key)

    def row(self, index):
        if self._index != index:
            self._index = index
            self._row += 1
        return self._row


    def write(self, row, col, value, index=''):
        if index:
            self.sheet.write(row, 0, index)
        if not col:
            pass
        if not value:
            self.sheet.write_blank(row, col, value)
        elif type(value) == bool:
            self.sheet.write_boolean(row, col, value)
        elif isinstance(value, str):
            self.sheet.write_string(row, col, value)
        else:
            self.sheet.write_number(row, col, value)


class ReadingSheet(ExcelSheet):

    def __init__(self, sheet):
        super().__init__(sheet)
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
    """A mixin class to provide functions to flatten or 
       unflatten dict.
    """

    def get_file_path(self, original_path, ext):
        suffix = datetime.now().strftime('%Y%m%d%H%M%S')
        path = '{}_{}{}'.format(os.path.splitext(original_path)[0],
            suffix, ext)
        return path


    def serialize(self, dic, idx, pref=''):
        """Flatten dict. 
        """
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
                    yield Cell(f'{pref}{key}{HYPHEN}{str(0)}', idx, val)
            else:
                yield f'{pref}{key}', idx, val

    
    def parse_json(self, dic, pref='', group=MAIN):
        """Return a pair of sheet name and column name.
        """
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
                    yield group, f'{pref}{key}{HYPHEN}{str(0)}'
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
    """Write data in json file to worksheets in Excel.
    """

    def __init__(self, json_file):
        json_file = os.path.abspath(json_file)
        self.excel_file = self.get_file_path(json_file, '.xlsx')
        self.json = ReadJson(json_file)
        self.excel_format = {}
        self.sheets = None
       

    def get_excel_format(self):
        """Create dict {column name: sheet name}
        """
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
        """Create dict {sh_name: WritingSheet instance}
           If a key is equal to sh_name directory connected to HYPHEN,
           the key is empty object list.
        """
        self.sheets = {}
        # Change sh_name if key is sh_name directory connected to HYPHEN 
        for sh_name in set(self.excel_format.values()):
            for key in self.excel_format.keys():
                if key == f'{sh_name}-0':
                    self.excel_format[key] = sh_name
        for sh_name in self.excel_format.values():
            if sh_name not in self.sheets:
                self.sheets[sh_name] = WritingSheet(
                    wb.add_worksheet(sh_name),
                    [key for key, val in self.excel_format.items() \
                        if val == sh_name and key != f'{sh_name}-0']
                )

  
    def write(self, sh_name, cell):
        """Get column and row numbers to write value to worksheet.
        """
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
                    self.write(sh_name, cell)


class FromExcel(Convert):
    """Read data on worksheets in Excel to output to json file.
       It is necessary to use Excel file created using ToExcel instance.
    """

    def __init__(self, excel_file):
        self.excel_file = os.path.abspath(excel_file)
        self.sheets = None
        
        
    def set_sheets(self, wb):
        self.sheets = tuple(ReadingSheet(sh) for sh \
            in wb if sh.cell(row=2, column=1).value)
   

    def convert(self, indent=None):
        wb = openpyxl.load_workbook(self.excel_file)
        self.set_sheets(wb)
        self.output(
            (record for record in self.read()),
            indent,
            self.get_file_path(self.excel_file, '.json')
        )
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
    

    def output(self, records, indent, json_file):
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(
                list(records), 
                f, 
                ensure_ascii=False,
                indent=indent
            )


if __name__ == '__main__':
    # test_dic1 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3]}
    # test_dic2 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3], 'e': [{'f': 5, 'g': 6}, {'f': 100, 'g': 120}]}
    # test_dic3 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3], 'e': [{'f': 5, 'g': 6, 'h': [89, 56, 23]}, {'f': 100, 'g': 120, 'h': [70, 56, 20]}]}
    # test_dic4 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [[1, 2, 3], [4, 5, 6]]}
    # test_dic5 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [[], []]}
    # test_dic6 = {'c': {'a': 2, 'b': {'x': 5, 'y': 10}, 'e': [10, 20, 30]}}

    # path = r"C:\Users\kanae\OneDrive\myDevelopment\JsonExcel\test_data\test5.json"
    to_excel = ToExcel('database.json')
    # to_excel = ToExcel(path)
    # print(to_excel.excel_file)
    to_excel.convert_all()
    # print({k : v for k, v in converter.serialize(test_dic3)})
    # path = r"C:\Users\kanae\OneDrive\myDevelopment\JsonExcel\database_20201114203603.xlsx"
    # path = r"C:\Users\kanae\OneDrive\myDevelopment\JsonExcel\dict10_20201116220255.xlsx"
    # from_excel = FromExcel(path)
    # from_excel.convert()
