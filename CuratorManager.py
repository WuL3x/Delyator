# import os
#
# from FileManager import FileManager
#
#
# class CuratorManager(FileManager):
#     def __init__(self, curator, manager, folder_name, file_name):
#         super().__init__(folder_name, file_name)
#         self.curator = curator
#         self.manager = manager
#
#     def create_curator_folder(self):
#         sanitize_filename = self.sanitize_filename()
#         curator_folder_path = os.path.join(self.folder_name, sanitize_filename)
#         os.makedirs(curator_folder_path, exist_ok=True)
#         return curator_folder_path
