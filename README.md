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
  
  *class* FromExcel(path)
  
   * Exported data in Excel to JSON file. 
   * The Excel must be the file output with convert method of ToExcel class.  
   
   ```bash
   from jsonexcel import FromExcel
   
   from_excel = FromExcel(path)                               # path: Excel file path
   from_excel.convert()                                       # Export data to JSON file
   from_excel(
       indent=4,                                              # If you need indent on JSON file, specify number.
       replacement={'name': 'my_name, 'age': 'my_age'}        # If you need to change key name, specify dict({key before: key after, ...})
    )
   
   ```
   
   
  # Note
  
   * If hyphens(-) or dots(.) are found in keys in a JSON file, they are replaced with underbar(\_) before export to Excel file.     
  
