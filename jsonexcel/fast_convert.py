from pyexcelerate import Workbook
from convert import ToExcel, WritingSheet


class FastWritingSheet(WritingSheet):

    URL = r"(https?|ftp)(:\/\/[-_\.!~*\'()a-zA-Z0-9;\/?:\@&=\+$,%#]+)"


    def set_keys(self, keys):
        self.keys = {}
        for col, key in enumerate(keys, 2):
            self.keys[key] = col
            self.write(1, col, key)


    def write(self, row, col, value, index=''):
        # print(row, col, value, index)
        try:
            if index:
                self.sheet.set_cell_value(row, 1, index)
            if not col:
                return
            if isinstance(value, list) or value is None:
                self.sheet.set_cell_value(row, col, '')
            else:
                # print(row, col, value, index)
                self.sheet.set_cell_value(row, col, value)
        except TypeError:
            # print(row, col, value, index)
            raise




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

    def write(self, cell):
        """Get column and row numbers to write value to worksheet.
        """
        sh_name = self.sheet_format.get(cell.key)
        sheet = self.sheets[sh_name]
        col = sheet.column(cell.key)
        row = sheet.row(cell.idx)
        # print(row, col, cell.value, cell.idx)
        sheet.write(row, col, cell.value, cell.idx)



    def output(self, records):
        excel_file = self.get_file_path(self.json.file, '.xlsx')
        wb = Workbook()
        self.set_sheets(wb)
        for record in records:
            for cell in record:
                self.write(cell)
        wb.save(excel_file)


if __name__ == '__main__':
    import time
    start = time.time()

    to_excel = FastToExcel('database.json')
    # to_excel.set_sheet_format()
    to_excel.convert()
    # print(to_excel.sheet_format)
    print(f'It took {time.time() - start} seconds')

