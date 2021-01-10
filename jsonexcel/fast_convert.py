from pyexcelerate import Workbook
from jsonexcel import ToExcel, WritingSheet


class FastWritingSheet(WritingSheet):

    
    def set_keys(self, keys):
        self.keys = {}
        for col, key in enumerate(keys, 2):
            self.keys[key] = col
            self.write(1, col, key)


    def write(self, row, col, value, index=''):
        if index:
            self.sheet.set_cell_value(row, 1, index)
        if not col:
            return
        if isinstance(value, list) or value is None:
            self.sheet.set_cell_value(row, col, None)
        else:
            self.sheet.set_cell_value(row, col, value)
        

class FastToExcel(ToExcel):


    def set_sheets(self, wb):
        """Create dict {sh_name: WritingSheet instance}
           If a key is equal to sh_name directory connected to HYPHEN,
           the key is empty object list.
        """
        self.sheets = {}
        # Change sh_name if key is sh_name connected to directory HYPHEN 
        for sh_name in set(self.sheet_format.values()):
            for key in self.sheet_format.keys():
                if key == f'{sh_name}-0':
                    self.sheet_format[key] = sh_name
        for sh_name in self.sheet_format.values():
            if sh_name not in self.sheets:
                self.sheets[sh_name] = FastWritingSheet(
                    wb.new_sheet(sh_name),
                    sorted(key for key, val in self.sheet_format.items() \
                        if val == sh_name and key != f'{sh_name}-0'),
                    row=1
                )


    def output(self, records):
        excel_file = self.get_file_path(self.json.file, '.xlsx')
        wb = Workbook()
        self.set_sheets(wb)
        for record in records:
            for cell in record:
                self.write(cell)
        wb.save(excel_file)
