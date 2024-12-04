import os


class FileManager:
    def __init__(self, file_path):
        self.file_name = file_path

    def create_output_folder(self, file_path):
        self.file_name = os.path.basename(file_path)
        folder_name = os.path.splitext(self.file_name)[0]
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
        os.makedirs(folder_path, exist_ok=True)
        return folder_path
    @staticmethod
    def create_file(parent_folder, name):
        sanitized_filename = FileManager.sanitize_filename(name)
        return os.path.join(parent_folder, f'{sanitized_filename}.xlsx')
