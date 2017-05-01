# Excel functions
A set of custom-made modules which help use Excel.

### xls.py
This class helps find values in the old 2003 .xls Excel file format
```python
from xls import XlsTools

xl = XlsTools('file.xls')
# find cell with 'Description', returning cell 1 to the right
summary = xl.find('Description', 1).strip()
```

### xlsx.py
This class helps you preform common operations with the newer xlsx file format.
```python
from xlsx import XlsxTools

xlsx = XlsxTools()
contents = []
xlsx.create_document(contents, 'my tab name', 'file_out.xlsx')
```

### excel_com.py
This is a Python class which provides common Microsoft Excel functions using the Win32com API.

```python
from excel_com import Excel

e = Excel()
e.open(r'c:\original_file.xlsx', read_only=True)
# save with default settings
e.save_as(r'c:\new_file.xlsx', True)
e.close()
```
