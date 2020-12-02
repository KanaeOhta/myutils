# JsonExcel

Export JSON-format data to Excel like RDB or the exported data in Excel to JSON file.


# Requirements

* Python 3.8
* openpyxl
* xlsxwriter


# Environment

* Windows10


# Usage

 *class* ToExcel(path)
 
  * Export JSON-format data to Excel like RDB
  
  ```bash
  from jsonexcel import ToExcel
  
  to_excel = ToExcel(path)                                    # path: JSON file path
  to_excel.convert()                                          # When export all data
  to_excel.partial_convert(column name 1, column name 2, ...) # When export selected data

  ```
  
  
