from DeltaCalculation import deltaCalculation

# start_point = [5, 'C']
excel_file_path = "../Sample/Sample_Operations Summary Q1'24 Week 7 v4 version check - Copy.xlsx"

sheet_list = [
    [[9, "K"], "Total AMAT", "AMAT wk 6"],
    [[8, "I"], "SPG", "SPG wk6"],
    [[8, "I"], "AGS", "AGS wk6"],
    [[8, "I"], "Display", "Display wk6"],
    [[8, "I"], "Corporate", "Corporate wk6"]
]

for i in range(len(sheet_list)):
    sheet = sheet_list[i]
    deltaCalculation(excel_file_path=excel_file_path, start_point=sheet[0], sheet_name_current=sheet[1],
                     sheet_name_previous=sheet[2])
