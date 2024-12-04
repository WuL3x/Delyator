import os
from tkinter import Tk, Label, Button, filedialog, simpledialog

from openpyxl.reader.excel import load_workbook

from FileManager import FileManager as fm
from ExcelOperations import ExcelOperations


def choose_file():
    root = Tk()
    root.withdraw()  # скрываем главное окно
    file_path = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[("Excel files", "*.xlsx")])
    return file_path


def choose_columns()->tuple:
    def column_letter_to_number(letter):
        return ord(letter.upper()) - 65

    definition_column_letter = simpledialog.askstring(title="Выбор столбца",
                                                      prompt="Введите букву столбца для куратора:")
    genus_column_letter = simpledialog.askstring(title="Выбор столбца", prompt="Введите букву столбца для начальника:")

    definition_column_num = column_letter_to_number(definition_column_letter)
    genus_column_num = column_letter_to_number(genus_column_letter)

    return definition_column_num, genus_column_num


def deploy_tkinter(func1, func2):
    root = Tk()
    root.title("Делятор-v2")
    root.geometry("400x200")

    status_label = Label(root, text="Нажмите кнопку 'Обработать файл', для начала работы")
    status_label.pack(pady=10)

    process_button = Button(root, text='Обработать файл', command='')
    process_button.pack(pady=5)
    deploy_tkinter.mainloop()


def create_folder_file(sheet, def_column, gen_column, output_folder):
    for row in sheet.max_row:
        definition_value = sheet.cell(row=row, column=def_column + 1).value
        genus_value = sheet.cell(row=row, column=gen_column + 1).value

        if not definition_value or not genus_value:  # Если нет значения, то пропуск
            continue

        definition_folder = fm.create_folder(output_folder, definition_value)
        genus_file = fm.create_file(definition_folder, genus_value)
    return definition_folder, genus_file

def sort_group(sheet, folder_column, file_column):
    data = []
    for row in range(2, sheet.max_row+1):
        folder_value = sheet.cell(row=row,column=folder_column + 1).value
        file_value = sheet.cell(row=row, column=file_column + 1).value
        if folder_value and file_value:
            data.append((folder_value, file_value))

    return sorted(data, key=lambda x:x[1])

def create_def_gen_folder_file(sheet, folder_name, file_name, output_folder) -> list:
    data = sort_group(sheet,folder_name,file_name)
    current_folder = None
    folder_path = "Не найдено"
    created_path = []
    for folder_name, file_name in data:
        if folder_name != current_folder:
            current_folder = folder_name
            sanitized_folder_name = fm.sanitize_filename(folder_name)
            folder_path = fm.create_folder(output_folder, sanitized_folder_name)

        sanitized_file_name = fm.sanitize_filename(file_name) + '.xlsx'
        file_path = os.path.join(folder_path, sanitized_file_name)
        created_path.append({
            file_path:folder_path
        })
    return created_path

def main():
    file_path = choose_file()
    excel_operations = ExcelOperations(path=file_path)  # экземпляр объекта ExcelOperations
    file_manager = fm(file_path)

    def activate_main_sheet():
        main_wb = load_workbook(file_path)
        return main_wb.active

    output_folder = file_manager.create_output_folder(file_path=file_path)
    main_sheet = activate_main_sheet()

    definiton_columns, genus_column = choose_columns()

    #
    # for row in main_sheet.max_row:
    #     definition_value = main_sheet.cell(row=row, column=definiton_columns + 1).value
    #     genus_value = main_sheet.cell(row=row, column=genus_column + 1).value
    #
    #     if not definition_value or not genus_value:  # Если нет значения, то пропуск
    #         continue
    #
    #     definition_folder = fm.create_folder(output_folder, definition_value)
    #     genus_file = fm.create_file(definition_folder, genus_value)



