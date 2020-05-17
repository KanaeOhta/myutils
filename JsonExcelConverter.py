import argparse
import glob
import os
import re

from tqdm import tqdm

from texttools.xlsxhandler import ExcelMixin
from texttools.jsonhandler import JsonHandler
from texttools.utils import create_folder, get_file


def handle_commandline():
    parser = argparse.ArgumentParser()
    parser.add_argument('file',
        choices=FileCheck(),
        help='<json/xlsx file path or folder containing one or more json/xlsx files')
    args = parser.parse_args()
    return args.file


class FileCheck:
    def __contains__(self, val):
        if os.path.isfile(val) or os.path.isdir(val):
            return True

        
    def __iter__(self):
        return iter(('json/xlsx',))


class JsonToExcel(ExcelMixin):

    def __init__(self, json_file):
        self.file = json_file
      

    def convert(self, folder):
        json_handler = JsonHandler(self.file)
        js_dicts = json_handler.read(
            lambda x: (self.flatten(js) for js in x)
        )
        output_file = get_file(folder, json_handler.file)
        self.json_to_excel(js_dicts, output_file)


class ExcelToJson(ExcelMixin):

    def __init__(self, excel_file):
        self.file = excel_file


    def convert(self, folder):
        output_file = get_file(folder, self.file, ext='json')
        json_handler = JsonHandler(output_file)
        json_handler.output(self.excel_to_json(self.file))


def convert_file(converters, folder):
    with tqdm(total=len(converters), ncols=70) as pbar:
        for converter in converters:
            converter.convert(folder)
            pbar.update(1)


def main():
    # while Excel file is opend, its temp file which name starts with '~$' is made in the same dir.
    # Trying to open such temp file results in error. 
    path = handle_commandline()
    files = [path] if os.path.isfile(path) else \
        [file for file in glob.glob('{}/*'.format(path)) if \
            re.search('/*.(json|xlsx)', file) and not os.path.basename(file).startswith('~$')]
    converters = []
    for file in files:
        ext = os.path.splitext(file)[-1].lower()
        if ext == '.json':
            converters.append(JsonToExcel(file))
        elif ext == '.xlsx':
            converters.append(ExcelToJson(file))    
    if converters:
        folder = create_folder(False, path, 'convert')
        convert_file(converters, folder)
        print('Successfully created files into {}'.format(folder))
    else:
        print('Finished, no files to output.')
    


if __name__ == '__main__':
    main()
   
