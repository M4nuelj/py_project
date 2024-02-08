import pandas as pd
import numpy as np
import warnings

warnings.filterwarnings("ignore")

file_path = "C:/Users/manue/OneDrive/Escritorio/py_project/audit_excel_2024.xlsx"

path = "C:/Users/manue/OneDrive/Escritorio/py_project/"

def read_excel_file(file_path):
  
    excel_file = pd.ExcelFile(file_path)

    dataframes = {sheet_name: excel_file.parse(sheet_name) for sheet_name in excel_file.sheet_names}

    return dataframes




result = read_excel_file(file_path)

Templates = ['1A PRODUCT INSPECTION - EQUIPME',
              '1B DAILY INSPECTION SHEET- MENU',
              '2A DAILY INSPECTION SHEET BLEND',
              '2B DAILY INSPECTION SHEET (BLEN',
              '3 DAILY INSPECTION SHEET - MANU',
              '4 DAILY INSPECTION SHEET - REPA']

def clean_dictionary(result, Templates):
    cleaned_dict = {}
    for sheet in Templates:
        if sheet in result:
            cleaned_dict[sheet] = result[sheet]
    return cleaned_dict


result2 = clean_dictionary(result, Templates)

Temp_Columns_1A = ['Title Page_Inspection by', 'Title Page_Inspection date', 'Title Page_Processed Product_Product item code',    
                   'Title Page_Lbs. Processed', 
                   'Title Page_Lbs. Inspected']

Temp_Columns_1B = ['Title Page_Inspected by', 'Title Page_Inspection date', 'Title Page_Processed Product_Processed item code',    
                   'Title Page_Lbs. Processed', 
                   'Title Page_Lbs. Inspected']

Temp_Columns_2A = ['Title Page_Inspected by', 'Title Page_Inspection date', 'Title Page_Processed Product_Product item code',    
                   'Title Page_Lbs. Processed', 
                   'Title Page_Lbs. Inspected']

Temp_Columns_2B = ['Title Page_Inspection by', 'Title Page_Inspection date', 'Title Page_Processed Product_Product item code',    
                   'Title Page_Lbs. Processed', 
                   'Title Page_Lbs. Inspected']

Temp_Columns_3 = ['Title Page_Inspection by', 'Title Page_Inspection date', 'Title Page_Processed Product_Processed item code',    
                   'Title Page_Lbs. Processed', 
                   'Title Page_Lbs. Inspected']

Temp_Columns_4 = ['Title Page_Inspection by', 'Title Page_Inspection date', 'Title Page_Processed Product_Product item code',    
                   'Title Page_Lbs. Processed', 
                   'Title Page_Lbs. Inspected']

DFs = []

for key,values in result2.items():
    if key == Templates[0]:
        df_1A = result2[key][Temp_Columns_1A]
        df_1A.columns = ['Inspected_by', 'Date', 'Item_Code', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_1A)
    elif key == Templates[1]:
        df_1B = result2[key][Temp_Columns_1B]
        df_1B.columns = ['Inspected_by', 'Date', 'Item_Code', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_1B)
    elif key == Templates[2]:
        df_2A = result2[key][Temp_Columns_2A]
        df_2A.columns = ['Inspected_by', 'Date', 'Item_Code', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_2A)
    elif key == Templates[3]:
        df_2B = result2[key][Temp_Columns_2B]
        df_2B.columns = ['Inspected_by', 'Date', 'Item_Code', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_2B)
    elif key == Templates[4]:
        df_3 = result2[key][Temp_Columns_3]
        df_3.columns = ['Inspected_by', 'Date', 'Item_Code', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_3)
    else:
        df_4 = result2[key][Temp_Columns_4]
        df_4.columns = ['Inspected_by', 'Date', 'Item_Code', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_4)

for df in DFs:
    df.replace(' ', np.nan, inplace = True)
    df.dropna(inplace = True)
    df.loc[:,'Item_Code'] = df.loc[:,'Item_Code'].str.upper()
    df.loc[:,'Date'] = df.loc[:,'Date'].dt.strftime("%m/%d/%Y")

Total_Inspections = pd.concat(DFs, axis=0)

Total_Inspections.reset_index(drop = True, inplace = True)

try:
    Total_Inspections.to_excel(path + 'Total_Inspections.xlsx', index = False)
except PermissionError:
    print("The file 'Total_Inspections.xlsx' is open. Please close it and run the code again.")
else:
    print("File saved successfully.")
