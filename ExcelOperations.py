from copy import copy
from tkinter import Tk, filedialog, simpledialog
import logging

class ExcelOperations:
    def __init__(self, path: str):
        print(f'Инициализация: path={path}')
        if path.endswith('xlsx'):
            self.path = path
        else:
            raise FileExistsError(f"Файл {path} не найден или не поддерживается.")

    @staticmethod
    def apply_cell_styles(main_cell, new_cell):
        new_cell.font = copy(main_cell.font)
        new_cell.fill = copy(main_cell.fill)
        new_cell.border = copy(main_cell.border)
        new_cell.alignment = copy(main_cell.alignment)
        new_cell.number_format = copy(main_cell.number_format)
        new_cell.protection = copy(main_cell.protection)

    @staticmethod
    def create_header(source_sheet, target_sheet):

        for col_num, cell in enumerate(source_sheet[1], start=1):
            if isinstance(cell, tuple):
                cell = cell[0]
            new_cell = target_sheet.cell(row=1, column=col_num, value=cell.value)

            ExcelOperations.apply_cell_styles(cell, new_cell)
        logging.info("Заголовок успешно скопирован на новый лист с сохранением стилей")


    @staticmethod
    def copy_row(main_sheet, new_sheet, row, new_row_index):
        for col in range(1, main_sheet.max_column + 1):
            cell = main_sheet.cell(row=row, column=col)
            new_cell = new_sheet.cell(row=new_row_index, column=col, value=cell.value)
            if cell.has_style:
                ExcelOperations.apply_cell_styles(cell, new_cell)
    @staticmethod
    def delete_colm(sheet, *colm_number):
        for colm_index in sorted(colm_number, reverse=True):
            sheet.delete_cols(colm_index + 1)

    @staticmethod
    def size_dims(main_sheet):
        dims = {}
        for row in main_sheet.rows:
            for cell in row:
                if cell.value:
                    col_char = chr(64 + cell.column)
                    if cell.column < 3:
                        min_wid = 7
                    else:
                        min_wid = 17
                    dims[col_char] = max((dims.get(cell.column, min_wid), len(str(cell.value))))

        return dims

    @staticmethod
    def set_columns_width(sheet, dims:dict):
        for col, width in dims.items():
            sheet.column_dimensions[col].width = width


