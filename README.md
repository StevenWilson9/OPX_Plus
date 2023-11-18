## Introduction

OPX_Plus (OpenPyXL Plus) is an extension to [OpenPyXL](https://pypi.org/project/openpyxl/), a module for reading and writing .xlsx and similar files. **OPX_Plus** adds functions centred around creating regular reports based on a template file. You can log issues on the [GitHub page](https://github.com/StevenWilson9/OPX_Plus/issues).

### Has the Following Dependencies
```
import openpyxl, glob, webbrowser, warnings, csv, datetime, os, time
from copy import copy
from openpyxl.formula.translate import Translator
```

## Functions
### Files
```
- get_file_path(file_path, name_search, *if_missing_urls)
- check_mandatory_files(f_list)

- open_template(template_path: object, *date_cells: str, previousday=False)
- save_file(new_file_location, file_name, workbook, afterdatetext="", previousday=False)
- remove(file_path)

```
### Import Data
```
- open_vals_only_sheet(from_wb_path, sheet_id=0):
- paste_sheet_to_sheet(from_ws, to_ws, cell_range)
- paste_to_sheet(xlsx or csv file location, ws)
- paste_df_to_sheet(df, to_ws, include_header=False)
```

### Move Data
```
- copy_over_and_down_formulas(from_ws, to_ws, formula_cells)
- paste_cells_to_cells(from_ws, to_ws, from_cell_range, offset=(0, 0))
```

### Sub-Functions
```
- count_rows(ws)
- count_columns(ws)
- num_to_excel_col(n)
- excel_col_to_num(a)

- to_unix_time(year=datetime.datetime.now().year, month=datetime.datetime.now().month,
                 day=datetime.datetime.now().day, hour=datetime.datetime.now().hour,
                 minute=datetime.datetime.now().minute, second=datetime.datetime.now().second, add_3_zeros=False,
                 subtract=datetime.timedelta(0))
- date_ranger()
- worksheet_reset(sheet_nam, wb)
```

