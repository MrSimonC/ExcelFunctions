# Excel functions
This is a Python class which provides common Microsoft Excel functions using the Win32com API.

```python
import excel_com

e = excel_com.Excel()
e.open(r'c:\original_file.xlsx', read_only=True)
# save with default settings
e.save_as(r'c:\new_file.xlsx', True)
e.close()
```
