from copy import copy
from tkinter import Tk, filedialog, simpledialog


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
    def create_header(main_sheet, new_sheet):
        for col, header_cell in enumerate(main_sheet[1], start=1):
            new_sheet.cell(row=1, column=col, value=header_cell.value)
            ExcelOperations.apply_cell_styles(header_cell, new_sheet.cell(row=1, column=col))
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


