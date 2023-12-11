import os
import openpyxl
from openpyxl.styles import PatternFill, Border
from tqdm import tqdm

FILE_EXTENSION = '.xlsx'

# Удаления пустых строк
def remove_empty_rows(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
        if any(cell.value for cell in row):
            continue

        for cell in row:
            cell.border = Border()
            cell.fill = PatternFill()

# Обработка всех листов
def process_workbook(workbook):
    for sheet_name in workbook.sheetnames:
        remove_empty_rows(workbook[sheet_name])

# Конкретный файл
def process_file(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        process_workbook(workbook)
        
        workbook.save(file_path)

    except Exception as e:
        print(f"{file_path}: {e}")

# Список файлов
def files_in_folder(folder_path):
    # Выбор нужного расширения
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f)) and f.endswith(FILE_EXTENSION)]

    # Прогресс
    for file in tqdm(files, desc="Processing files", unit="file"):
        file_path = os.path.join(folder_path, file)

        process_file(file_path)

files_in_folder(r'C:\Users\Админ\Desktop\excel')