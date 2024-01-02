import pandas as pd
import os

# Identifying the files to merge
path = r'C:\Users\annie\Desktop\Python\Sample\test'
files = os.listdir(path)
files_xlsx = [f for f in files if f.endswith('.xlsx')]

# Reading the Excel files and concatenate them
mydf_list = [pd.read_excel(os.path.join(path, f)) for f in files_xlsx]
mydf = pd.concat(mydf_list)

# Writing the merged data frame to a new Excel file
myoutput_path = r'C:\Users\annie\Desktop\Python\Sample\test\merged_workbook.xlsx'
mydf.to_excel(myoutput_path, index=False, )

df.to_excel(w, sheet_name = new_sheet_name , index = False)