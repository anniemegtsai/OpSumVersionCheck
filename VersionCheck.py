import pandas as pd
from openpyxl.utils import get_column_letter, column_index_from_string
from DeltaCalculation import deltaCalculation

start_point = [5, 'C']
excel_file_path = '../Sample/Sample.xlsx'

sheet_list = [
	[start_point, "month1", "month2"],
	[start_point, "month3","month4"]
]

for i in range(len(sheet_list)):
	sheet = sheet_list[i]
	deltaCalculation(excel_file_path = excel_file_path, start_point = sheet[0], sheet_name_current =sheet[1], sheet_name_previous = sheet[2])