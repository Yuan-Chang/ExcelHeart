import math
from utils.Utils import get_file_name_from_file_path, delete_directory, create_directory
from openpyxl import Workbook
from utils.openpyxl_helper import draw_heart

intensity = 20000

output_folder = "output/"
result_file = output_folder + "result.xlsx"

delete_directory(output_folder)
create_directory(output_folder)

wb = Workbook()
draw_heart(wb, sheet_name="abc", with_text=True)

wb.save(result_file)
