# pip install xrld==1.2.0 
import os
import pandas as pd
import xlrd
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

#Set Directory with files
os.chdir('./xls')

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
df = pd.DataFrame.from_dict(data)
df.to_excel("Scraped_Data.xlsx",sheet_name="Sheet1")