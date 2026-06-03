from FolderReader import FolderReader
from ExcelReader import ExcelReader

import pandas as pd


class SearchEngine:

    def __init__(self, folder_path):

        self.folder_path = folder_path

        folder_reader = FolderReader(folder_path)

        self.files = (
            folder_reader.get_all_excel_files()
        )

    # ================= GET ALL COLUMNS =================

    def get_all_columns(self):

        columns = set()

        for file in self.files:

            try:

                reader = ExcelReader(file)

                for col in reader.get_columns():

                    columns.add(
                        str(col).strip()
                    )

            except:
                pass

        return sorted(columns)

    # ================= SEARCH =================

    def run_search(
        self,
        checks,
        progress_callback=None
    ):

        matched_dataframes = []

        summary_results = []

        total = len(self.files)

        for index, file in enumerate(self.files):

            try:

                reader = ExcelReader(file)

                # Validate columns

                if not all(
                    check["column"] in reader.df.columns
                    for check in checks
                ):

                    continue

                matched = reader.condition_match(
                    checks
                )

                if not matched.empty:

                    matched["Source_File"] = file

                    matched_dataframes.append(
                        matched
                    )

                    summary_results.append(
                        (
                            file,
                            len(matched)
                        )
                    )

            except Exception as e:

                summary_results.append(
                    (
                        file,
                        f"ERROR: {e}"
                    )
                )

            if progress_callback:

                progress_callback(
                    index + 1,
                    total
                )

        # ================= MERGE =================

        if matched_dataframes:

            final_df = pd.concat(
                matched_dataframes,
                ignore_index=True
            )

        else:

            final_df = pd.DataFrame()

        return summary_results, final_df

    # ================= EXPORT =================

    def export_results(
        self,
        dataframe,
        output_path
    ):

        dataframe.to_excel(
            output_path,
            index=False
        ) 