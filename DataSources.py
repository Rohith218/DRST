from FolderReader import FolderReader
from ExcelReader import ExcelReader
import pandas as pd


# ================= EXCEL WRAPPER (NO LOGIC CHANGE) =================

class ExcelDataSource:

    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.files = FolderReader(folder_path).get_all_excel_files()

    def get_all_columns(self):

        columns = set()

        for file in self.files:
            try:
                reader = ExcelReader(file)
                columns.update(reader.get_columns())
            except:
                pass

        return sorted(columns)

    def search(self, checks, progress_callback=None):

        matched_dataframes = []
        summary_results = []

        total = len(self.files)

        for index, file in enumerate(self.files):

            try:
                reader = ExcelReader(file)

                # ❗ DO NOT TOUCH EXISTING LOGIC
                if not all(c["column"] in reader.df.columns for c in checks):
                    continue

                matched = reader.condition_match(checks)

                if not matched.empty:
                    matched["Source_File"] = file
                    matched_dataframes.append(matched)
                    summary_results.append((file, len(matched)))

            except Exception as e:
                summary_results.append((file, f"ERROR: {e}"))

            if progress_callback:
                progress_callback(index + 1, total)

        final_df = (
            pd.concat(matched_dataframes, ignore_index=True)
            if matched_dataframes else pd.DataFrame()
        )

        return summary_results, final_df

    def export(self, df, path):
        df.to_excel(path, index=False)


# ================= DB STUB (ONLY STRUCTURE) =================

class DatabaseDataSource:

    def __init__(self):
        self.conn = None  # Oracle connection later

    def get_all_columns(self):
        return []

    def search(self, checks, progress_callback=None):
        # Placeholder only
        return [], pd.DataFrame()

    def export(self, df, path):
        df.to_excel(path, index=False)