import logging
import os


class FileManager:
    @staticmethod
    def create_output_folder(file_path):
        file_name = os.path.basename(file_path)
        folder_name = os.path.splitext(file_name)[0]
        output_folder = os.path.join(os.path.expanduser("~"), "Downloads", folder_name)
        os.makedirs(output_folder, exist_ok=True)
        return output_folder

    @staticmethod
    def sanitize_filename(name):
        if not name:
            raise ValueError("Имя файла не задано")
        return "".join(c for c in str(name) if c.isalnum() or c in (" ", '_', ".")).rstrip()

    @staticmethod
    def create_folder(folder_name, name):
        sanitized_filename = FileManager.sanitize_filename(name)
        folder_path = os.path.join(folder_name, sanitized_filename)

        logging.info(f"Пытаемся создать папку по пути: {folder_path}")

        os.makedirs(folder_path, exist_ok=True)
        logging.info(f"Папка '{folder_path}' успешно создана или уже существует.")

        return folder_path

    @staticmethod
    def create_file(parent_folder, name, wb):

        sanitized_filename = FileManager.sanitize_filename(name)

        file_path = os.path.join(parent_folder, f'{sanitized_filename}.xlsx')

        if not os.path.exists(parent_folder):
            os.makedirs(parent_folder)
            logging.info(f"Создана папка: {parent_folder}")

        if os.path.exists(file_path):
            logging.info(f"Файл {file_path} уже существует.")
        else:
            wb.save(file_path)
            logging.info(f"Создан файл: {file_path}")

        return file_path