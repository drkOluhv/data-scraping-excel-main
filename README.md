## 1. Overview
Parse old .xls Excel files into a final dataframe summarising values from each file using the xlrd library. Now superceeded with .xls and removal of support for .xlsx files using openpxl or other python libraries.

## 1.1 Import Packages

```py
import os
import pandas as pd
import xlrd
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True
#Set Directory with files
os.chdir('./xls')
```

## 1.2 Sheets in the Workbook

Reading data from the workbook which can have multiple sheets requires specifying the worksheet within the workbbok which can be called by:

- Index: the index of the sheet starting with the first sheet at zero `sheet_by_index(0)`
- Name: str of the name of the sheet `sheet_by_name('Sheet1')`

```py
#To get sheet by index (starts with 0)
ws = wb.sheet_by_index(0)
#To get sheet by name 
ws = wb.sheet_by_name('Sheet1')
```

## 1.3 Values in the Sheet

Reading the data from a specific cell range by row and or column value requires an index argument

```py
#get list of values in row
ws.row_values(1)
#get list of values in column
ws.col_values(0)
```

## 1.4 Using Columnar Keys

When columns have attributes defining the values tabulated using the pre-defined attributes allows to map indexes to attributes

```py
#get the index of the attribute specified, e.g. First Name
attributes = ws.col_values(0)
attribute_index = attributes.index('First Name')

#get the neighboring value output of the attribute, e.g. Grace Hopper
values = ws.col_values(1)
values[attribute_index]
```

# 2. Functions

## 2.1 Get Value Function 

Using the xlrd package to open the workwork and find the `row_values` and `column_values` as attributes within a function allows to loop though the excel files.

The function to grab the attribute value can be completed where the value is adjacent to the pre-defined attribute

```py
def get_value(worksheet, attribute_column, attribute_name):
    attributes = worksheet.col_values(attribute_column)
    if attribute_name in attributes:
        attribute_index = attributes.index(attribute_name)
        #assume value is in the adjacent column where attribute is stored
        values = worksheet.col_values(attribute_column+1)
        value = values[attribute_index]
        return value
    else:
        return None
```

## 2.2 Loop Excel Files & Collect Values

Using `os.walk` to parse through the directory opening each workbook for the fist Sheet and grabbing the values assocaited with the pre-defined attributes and appending to the dictionary allows for scaping of all excel files for the values specified by the attributes.

```py
for root, dirs, files in os.walk('.'):
    attributes = ['First Name', 'Last Name', 'Sex','City','State']
    #initialized dictionary, create empty list for attributes with dict comprehension
    data = {attribute: [] for attribute in attributes}
    #append a key:value for File, will use this as unique identifier/index
    data.update({"File": []})
    for file in files:
        wb = xlrd.open_workbook(file)
        ws = wb.sheet_by_index(0)
        data['File'].append(file)
        for attribute in attributes:
            data[attribute].append(get_value(ws,0,attribute))
data
```

## 2.3 Resulting Summary

After creating the dictionary with all the attribute values in each Excel file the dictionary can be converted to a dataframe  with pandas and exported to a summarised Excel file summarising the information.

```py
df = pd.DataFrame.from_dict(data)
df.to_excel("Scraped_Data.xlsx",sheet_name="Sheet1")
```