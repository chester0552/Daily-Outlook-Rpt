from pathlib import Path
import time
import openpyxl
import warnings
import shutil
import os

# Suppress the openpyxl warning about missing default style
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles")

t = time.localtime()
timestamp = time.strftime('%Y-%m-%d', t)

output_dir = Path.cwd() / "Output"
output_dir.mkdir(parents=True, exist_ok=True)

def get_todays_excel_files():
    result = []
    for excel_file in Path(output_dir).glob('*xlsx'):
        file_stat = excel_file.stat()
        modification_time = time.strftime('%Y-%m-%d', time.localtime(file_stat.st_mtime))

        if modification_time == timestamp:
            result.append(excel_file)

    return result

todays_files = get_todays_excel_files()


def process_excel_file():

    for filename in todays_files:
        wb = openpyxl.load_workbook(filename)
        sheet = wb['Sheet0']
        row_number_to_delete = 11  # Replace this with the actual row number you want to delete
        row_number_to_delete2 = 10
        row_number_to_delete3 = 2
        row_number_to_delete4 = 1
        sheet.delete_rows(row_number_to_delete, 1)
        sheet.delete_rows(row_number_to_delete2, 1)
        sheet.delete_rows(row_number_to_delete3, 1)
        sheet.delete_rows(row_number_to_delete4, 1)
        wb.save(filename)

    return todays_files


completed_excels = process_excel_file()

final_dir = Path.cwd() / "Idle Code Report"
final_dir.mkdir(parents=True, exist_ok=True)


for source in completed_excels:
    destination = final_dir / source.name
    shutil.move(source, destination)
