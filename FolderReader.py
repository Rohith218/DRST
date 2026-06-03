import os


class FolderReader:

    def __init__(self, root_directory=None):

        self.root_directory = root_directory or os.getcwd()

    def get_all_excel_files(self):

        excel_files = []

        for root, dirs, files in os.walk(self.root_directory):

            for file in files:

                # Ignore Excel temp files
                if file.startswith("~$"):
                    continue

                if file.lower().endswith((".xlsx", ".xlsm", ".xls")):

                    full_path = os.path.join(root, file)

                    excel_files.append(full_path)

        return excel_files