import os
from concurrent.futures import ThreadPoolExecutor
from tkinter import Tk, Label, Button, filedialog, simpledialog
import logging
from openpyxl import load_workbook, Workbook

from FileManager import FileManager as fm
from ExcelOperations import ExcelOperations

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


root = Tk()
root.withdraw()

def choose_file():
    logging.info("Запрос на выбор файла Excel")
    file_path = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[("Excel files", "*.xlsx")])
    return file_path

def choose_columns() -> tuple:
    logging.info("Запрос на выбор столбцов для куратора и начальника")
    def column_letter_to_number(letter):
        return ord(letter.upper()) - 65

    genus_column_letter = simpledialog.askstring(title="Выбор столбца", prompt="Введите букву столбца для куратора:")
    species_column_letter = simpledialog.askstring(title="Выбор столбца", prompt="Введите букву столбца для начальника:")

    genus_column_num = column_letter_to_number(genus_column_letter)
    species_column_num = column_letter_to_number(species_column_letter)
    logging.info(f"Выбраны столбцы: Куратор - {genus_column_num}, Начальник - {species_column_num}")
    return genus_column_num, species_column_num

def deploy_tkinter(func1):
    logging.info("Запуск графического интерфейса")
    root.deiconify()
    root.title("Делятор-v2")
    root.geometry("400x200")

    status_label = Label(root, text="Нажмите кнопку 'Обработать файл', для начала работы")
    status_label.pack(pady=10)

    def on_process_button_click():
        logging.info("Обработка файла началась")
        func1()

    process_button = Button(root, text='Обработать файл', command=on_process_button_click)
    process_button.pack(pady=5)
    return root

def sort_group(file_path, sheet, folder_column, file_column):
    logging.info(f"Начало сортировки данных по столбцу {file_column}")
    main_workbook = sheet.parent
    new_sheet_title = f'sorted_{sheet.title}'
    if new_sheet_title in main_workbook.sheetnames:
        logging.info(f"Лист '{new_sheet_title}' существует. Удаляем его.")
        del main_workbook[new_sheet_title]

    new_sheet = main_workbook.create_sheet(new_sheet_title)

    rows_styles = []
    for row in sheet.iter_rows():
        row_data = [(cell.value, cell) for cell in row]
        rows_styles.append(row_data)

    header = rows_styles[0]
    data = rows_styles[1:]
    sorted_data = sorted(data, key=lambda x: (
        x[file_column][0] if x[file_column][0] is not None else "",
        x[folder_column][0] if x[folder_column][0] is not None else ""
    ))

    for i, row_data in enumerate([header] + sorted_data, start=1):
        for j, (value, cell) in enumerate(row_data, start=1):
            new_cell = new_sheet.cell(row=i, column=j)
            new_cell.value = value
            if cell.has_style:
                ExcelOperations.apply_cell_styles(cell, new_cell)
    main_workbook.save(file_path)
    logging.info(f"Сортировка завершена. Данные сохранены в новый лист '{new_sheet_title}'")

    return new_sheet

def process_group(sheet, folder_name, files, output_folder):
    logging.info(f"Начинается обработка папки {folder_name}")

    sanitized_folder_name = fm.sanitize_filename(folder_name)
    folder_path = os.path.join(output_folder, sanitized_folder_name)

    os.makedirs(folder_path, exist_ok=True)
    logging.info(f"Папка создана или уже существует: {folder_path}")

    for file_name, rows in files.items():
        sanitized_file_name = fm.sanitize_filename(file_name)
        file_path = os.path.join(folder_path, f'{sanitized_file_name}.xlsx')

        new_wb = Workbook()
        new_sheet = new_wb.active
        dims = ExcelOperations.size_dims(sheet)
        ExcelOperations.create_header(sheet, new_sheet)
        ExcelOperations.set_columns_width(new_sheet, dims)

        row_index = 2

        for row in rows:
            ExcelOperations.copy_row(sheet, new_sheet, row[0].row, row_index)
            row_index += 1


        new_wb.save(file_path)
        logging.info(f"Данные для начальника {file_name} сохранены в файл: {file_path}")

def process_multi(sheet, gen_column, spc_column, output_folder):
    logging.info("Группировка данных по начальнику и куратору")

    grouped_data = {}
    for row in sheet.iter_rows(min_row=2):
        folder_name = row[gen_column].value
        file_name = row[spc_column].value
        if not folder_name or not file_name:
            continue

        if folder_name not in grouped_data:
            grouped_data[folder_name] = {}

        if file_name not in grouped_data[folder_name]:
            grouped_data[folder_name][file_name] = []

        grouped_data[folder_name][file_name].append(row)

    tasks = [(sheet, folder_name, files, output_folder) for folder_name, files in grouped_data.items()]

    with ThreadPoolExecutor() as executor:
        executor.map(lambda p: process_group(*p), tasks)

def main():
    file_path = choose_file()

    def activate_main_sheet():
        main_wb = load_workbook(file_path)
        return main_wb.active

    output_folder = fm.create_output_folder(file_path=file_path)
    logging.info(f"Создана выходная папка: {output_folder}")
    main_sheet = activate_main_sheet()

    genus_column_num, species_column_num = choose_columns()
    sorted_sheet = sort_group(file_path, main_sheet, genus_column_num, species_column_num)

    process_multi(sorted_sheet, genus_column_num, species_column_num, output_folder)

if __name__ == "__main__":
    root = deploy_tkinter(main)
    root.mainloop()
