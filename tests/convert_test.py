import glob
import json
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, main

from jsonexcel import ToExcel, FromExcel, ExcelSheet, JsonFile


TEST_DIR = None


def setUpModule():
    global TEST_DIR
    TEST_DIR = TemporaryDirectory()
    test_path = Path(TEST_DIR.name)
    for key, val in globals().items():
        if key.startswith(('dic', 'replacedic')):
            file_path = os.path.join(test_path, f'{key}.json')
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump([val], f, ensure_ascii=False)
                f.flush()
    for json_file in glob.glob(f'{TEST_DIR.name}/*.json'):
            to_excel = ToExcel(json_file)
            to_excel.convert()
    for excel_file in glob.glob(f'{TEST_DIR.name}/*.xlsx'):
        from_excel = FromExcel(excel_file)
        from_excel.convert(
            replacement={'a_a': 'a-a', 'c_c': 'c.c', 'd_d': 'd-d'})
   

def tearDownModule():
    TEST_DIR.cleanup()


class ToExceTestCase(TestCase):

    def get_test_files(self, file_name):
        for json_file in glob.glob(f'{TEST_DIR.name}/*.json'):
            if '_' not in (base := os.path.splitext(os.path.basename(json_file))[0]) \
                    and base.startswith(file_name):
                str_num = base.replace(file_name, '')
                yield json_file, str_num


    def test_serialize(self):
        for json_file, str_num in self.get_test_files('dic'):
            serialized = globals()[f'serialized{str_num}']
            to_excel = ToExcel(json_file)
            for record in to_excel.get_records():
                result = {(cell.key, cell.idx): cell.value for cell in record}
            with self.subTest(result):
                self.assertEqual(serialized, result)


    def test_parse_json(self):
        for json_file, str_num in self.get_test_files('dic'):
            parsed = globals()[f'parsed{str_num}']
            to_excel = ToExcel(json_file)
            to_excel.set_sheet_format()
            with self.subTest(to_excel.sheet_format):
                self.assertEqual(parsed, to_excel.sheet_format)

    
    def test_selected_serialize(self):
        for json_file, str_num in self.get_test_files('dic'):
            if item := selected_test.get(int(str_num)):
                keys, expect = item
                to_excel = ToExcel(json_file)
                for record in to_excel.get_selected_records(keys):
                    result = {(cell.key, cell.idx): cell.value for cell in record}
                with self.subTest(result):
                    self.assertEqual(expect, result)



class FromExcelTestCase(TestCase):

    def get_test_files(self, prefix):
        for json_file in glob.glob(f'{TEST_DIR.name}/*.json'):
            if '_' in (file_name := os.path.basename(json_file)) \
                    and file_name.startswith(prefix):
                obj_name = file_name.split('_')[0]
                yield json_file, obj_name


    def test_deserialize(self):
        for json_file, obj_name in self.get_test_files('dic'):
            expected = globals()[obj_name]
            read_json = iter(JsonFile(json_file))
            dic = next(read_json)
            with self.subTest(dic):
                self.assertEqual(expected, dic)

    
dic1 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': [1, 2, 3]}
dic2 = {'e': [{'f': 5, 'g': 6}, {'f': 100, 'g': 120}]}
dic3 = {'e': [{'h': [89, 56]}, {'h': [70, 56]}]}
dic4 = {'d': [[1, 2, 3], [4, 5, 6]]}
dic5 = {'a': [], 'b': [[], []]}
dic6 = {'a': 1, 'c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd': []}
dic7 = {'e': [{'f': 5, 'g': 6}, {'f': 100, 'g': 120}], 'h': [{'i': 5, 'j': 6}, {'i': 100, 'j': 120}]}
dic8 = {'h': [{'i': {'k': [5, 6], 'l': 100}, 'j': 6}, {'i': {'k': [7, 8], 'l': 150}, 'j': 120}]}
dic9 = {'c': [{'f': 5, 'g': [{'h': 100, 'i': 120}, {'h': 200, 'i': 220}]}, {'f': 7, 'g': [{'h': 150, 'i': 180}]}]}
dic10 = {'c': [{'f': 5, 'g': [{'h': 100, 'i': {'j': 15, 'k':16}}, {'h': 200, 'i': {'j': 17, 'k':18}}]}, {'f': 7, 'g': [{'h': 150, 'i': {'j': 19, 'k':20}}]}]}
dic11 = {'a-a': 1, 'c.c': {'a': 2, 'b': {'x': 5, 'y': 10}}, 'd-d': [1, 2, 3]}
dic12 = {'d': [[{'h': 5, 'i': 10}, {'h': 50, 'i': 1000}], [{'h': 50, 'i': 30}, {'h': 80, 'i': 100}]]}

serialized1 = {('a', '1'): 1, ('c.a', '1'): 2, ('c.b.x', '1'): 5, ('c.b.y', '1'): 10, ('d-0', '1'): 1, ('d-1', '1'): 2, ('d-2', '1'): 3}
serialized2 = {('e.f', '1-0'): 5, ('e.g', '1-0'): 6, ('e.f', '1-1'): 100, ('e.g', '1-1'): 120}
serialized3 = {('e.h-0', '1-0'): 89, ('e.h-1', '1-0'): 56, ('e.h-0', '1-1'): 70, ('e.h-1', '1-1'): 56}
serialized4 = {('d-0-0', '1'): 1, ('d-0-1', '1'): 2, ('d-0-2', '1'): 3, ('d-1-0', '1'): 4, ('d-1-1', '1'): 5, ('d-1-2', '1'): 6}
serialized5 = {('a-0', '1'): [], ('b-0-0', '1'): [], ('b-1-0', '1'): []}
serialized6 = {('a', '1'): 1, ('c.a', '1'): 2, ('c.b.x', '1'): 5, ('c.b.y', '1'): 10, ('d-0', '1'): []}
serialized7 = {('e.f', '1-0'): 5, ('e.g', '1-0'): 6, ('e.f', '1-1'): 100, ('e.g', '1-1'): 120, 
    ('h.i', '1-0'): 5, ('h.j', '1-0'): 6, ('h.i', '1-1'): 100, ('h.j', '1-1'): 120}
serialized8 = {('h.i.k-0', '1-0'): 5, ('h.i.k-1', '1-0'): 6, ('h.i.l', '1-0'):100, ('h.j', '1-0'): 6,
    ('h.i.k-0', '1-1'): 7, ('h.i.k-1', '1-1'): 8, ('h.i.l', '1-1'): 150, ('h.j', '1-1'): 120}
serialized9 = {('c.f', '1-0'): 5, ('c.g.h', '1-0-0'): 100, ('c.g.i', '1-0-0'): 120, ('c.g.h', '1-0-1'): 200, ('c.g.i', '1-0-1'): 220,
    ('c.f', '1-1'): 7, ('c.g.h', '1-1-0'): 150, ('c.g.i', '1-1-0'): 180}
serialized10 = {('c.f', '1-0'): 5, ('c.g.h', '1-0-0'): 100, ('c.g.i.j', '1-0-0'): 15, ('c.g.i.k', '1-0-0'): 16, 
    ('c.g.h', '1-0-1'): 200, ('c.g.i.j', '1-0-1'): 17, ('c.g.i.k', '1-0-1'): 18,
    ('c.f', '1-1'): 7, ('c.g.h', '1-1-0'): 150, ('c.g.i.j', '1-1-0'): 19, ('c.g.i.k', '1-1-0'): 20} 
serialized11 = {('a_a', '1'): 1, ('c_c.a', '1'): 2, ('c_c.b.x', '1'): 5, 
    ('c_c.b.y', '1'): 10, ('d_d-0', '1'): 1, ('d_d-1', '1'): 2, ('d_d-2', '1'): 3}
serialized12 = {('d-0.h', '1-0'): 5, ('d-0.i', '1-0'): 10, ('d-0.h', '1-1'): 50, ('d-0.i', '1-1'): 1000,
    ('d-1.h', '1-0'): 50, ('d-1.i', '1-0'): 30, ('d-1.h', '1-1'): 80, ('d-1.i', '1-1'): 100}

parsed1 = {'a': ExcelSheet.MAIN, 'c.a': ExcelSheet.MAIN, 'c.b.x': ExcelSheet.MAIN, 'c.b.y': ExcelSheet.MAIN, 'd-0': ExcelSheet.MAIN,
    'd-1': ExcelSheet.MAIN, 'd-2': ExcelSheet.MAIN}
parsed2 = {'e.f': 'e', 'e.g': 'e'}
parsed3 = {'e.h-0': 'e', 'e.h-1': 'e'}
parsed4 = {'d-0-0': ExcelSheet.MAIN, 'd-0-1': ExcelSheet.MAIN, 'd-0-2': ExcelSheet.MAIN, 'd-1-0': ExcelSheet.MAIN, 'd-1-1': ExcelSheet.MAIN, 'd-1-2': ExcelSheet.MAIN}
parsed5 = {'a-0': ExcelSheet.MAIN, 'b-0-0': ExcelSheet.MAIN, 'b-1-0': ExcelSheet.MAIN}
parsed6 = {'a': ExcelSheet.MAIN, 'c.a': ExcelSheet.MAIN, 'c.b.x': ExcelSheet.MAIN, 'c.b.y': ExcelSheet.MAIN, 'd-0': ExcelSheet.MAIN}
parsed7 = {'e.f': 'e', 'e.g': 'e', 'h.i': 'h', 'h.j': 'h'}
parsed8 = {'h.i.k-0': 'h', 'h.i.k-1': 'h', 'h.i.l': 'h', 'h.j': 'h'}
parsed9 = {'c.f': 'c', 'c.g.h': 'c.g', 'c.g.i': 'c.g'}
parsed10 = {'c.f': 'c', 'c.g.h': 'c.g', 'c.g.i.j': 'c.g', 'c.g.i.k': 'c.g'} 
parsed11 = {'a_a': ExcelSheet.MAIN, 'c_c.a': ExcelSheet.MAIN, 'c_c.b.x': ExcelSheet.MAIN, 
    'c_c.b.y': ExcelSheet.MAIN, 'd_d-0': ExcelSheet.MAIN, 'd_d-1': ExcelSheet.MAIN, 'd_d-2': ExcelSheet.MAIN}
parsed12 = {'d-0.h': 'd-0', 'd-0.i': 'd-0', 'd-1.h': 'd-1', 'd-1.i': 'd-1'}

selected_test = {
    # value is a tuple of selected keys and expected value
    1: (('a', 'c.b.x'), {('a', '1'): 1, ('c.b.x', '1'): 5}),
    2: (('e.g',), {('e.g', '1-0'): 6, ('e.g', '1-1'): 120}),
    7: (('h.i', 'h.j'), {('h.i', '1-0'): 5, ('h.j', '1-0'): 6, ('h.i', '1-1'): 100, ('h.j', '1-1'): 120}),
    9: (('c.f', 'c.g.i'), {('c.f', '1-0'): 5, ('c.g.i', '1-0-0'): 120, ('c.g.i', '1-0-1'): 220, ('c.f', '1-1'): 7, ('c.g.i', '1-1-0'): 180}),
    10: (('c.g.i.k',), {('c.g.i.k', '1-0-0'): 16, ('c.g.i.k', '1-0-1'): 18, ('c.g.i.k', '1-1-0'): 20})
}


if __name__ == '__main__':
    main()