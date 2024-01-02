from OpSumVersionCheck.DeltaCalculation import delta_calculation
import Utils.Utils as utils
from Utils.modules import Sheet
import pandas as pd

# start_point = [5, 'C']
# excel_file_path = "../../Sample/Sample_Operations Summary Q1'24 Week 7 v4 version check - Copy.xlsx"
# excel_file_path = "../../Sample/Sample - Copy.xlsx"

sample_file = "../../Sample/Sample - Copy.xlsx"
output_folder = "../../Sample/output/"

sample_copy_file = output_folder + utils.get_file_name_from_file_path(sample_file)
sample_copy_file = utils.replace_extension(sample_copy_file, "xlsx")

result_file = output_folder + "result.xlsx"

# Kill excel sheet
utils.quit_excel()

utils.delete_directory(output_folder)
utils.create_directory(output_folder)

# Create a sample copy
utils.create_file_copy(sample_file, sample_copy_file)

# Prepare the sheet
# utils.create_values_only_excel_file(input_file=sample_copy_file, output_file=result_file)

all_sheet_list = utils.load_all_sheets(excel_file_path=sample_copy_file)

# sheet_list = [
#     [[9, "K"], "Total AMAT", "AMAT wk 6"],
#     [[8, "I"], "SPG", "SPG wk6"],
#     [[8, "I"], "AGS", "AGS wk6"],
#     [[8, "I"], "Display", "Display wk6"],
#     [[8, "I"], "Corporate", "Corporate wk6"]
# ]

sheet_list = [
    [[5, "C"], "month1", "month2"],
    [[5, "C"], "month3", "month4"]
]

for i in range(len(sheet_list)):
    sheet = sheet_list[i]
    sheet_name_delta = sheet[1] + " delta"
    result_data_frame = delta_calculation(excel_file_path=sample_copy_file, start_point=sheet[0],
                                          sheet_name_current=sheet[1],
                                          sheet_name_previous=sheet[2])
    all_sheet_list.append(Sheet(name=sheet_name_delta, data_frame=result_data_frame, start_point=sheet[0]))

utils.write_sheets_to_excel(result_file, all_sheet_list)

# Write the data into the Delta sheet
#
# num = random.randint(5,10000)
# outputFileName = f"../../Sample/Sample{num}.xlsx"
# with pd.ExcelWriter(outputFileName, engine='openpyxl', mode='a') as writer:
#     result_df = pd.DataFrame(delta_data)
#     result_df.to_excel(writer, sheet_name=sheet_name_delta, index=False, startrow=header_index)