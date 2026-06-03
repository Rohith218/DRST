import pandas as pd
import os


class ExcelReader:

    def __init__(self, file_path):

        self.file_path = file_path

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"{file_path} not found")

        # Read Excel safely
        self.df = pd.read_excel(
            file_path,
            dtype=str,
            engine="openpyxl"
        ).fillna("")

        # Normalize column names (VERY IMPORTANT)
        self.df.columns = [
            str(col).strip()
            for col in self.df.columns
        ]

    # ================= GET COLUMNS =================

    def get_columns(self):
        return list(self.df.columns)

    # ================= VALIDATE =================

    def validate_column(self, column_name):
        return column_name in self.df.columns

    # ================= CONDITION MATCH =================

    def condition_match(self, checks):

        if not checks:
            return pd.DataFrame()

        condition = pd.Series([True] * len(self.df))

        for check in checks:

            column = check["column"].strip()
            operator = check["operator"].strip().lower()
            value = str(check["value"]).strip()

            if column not in self.df.columns:
                condition &= False
                continue

            series = self.df[column].astype(str).str.strip()

            current = pd.Series([False] * len(self.df))

            # ================= EQUALS =================
            if operator == "equals":
                current = series.str.lower() == value.lower()

            # ================= NOT EQUALS =================
            elif operator == "not equals":
                current = series.str.lower() != value.lower()

            # ================= GREATER THAN =================
            elif operator == "greater than":
                numeric_series = pd.to_numeric(series, errors="coerce")

                try:
                    numeric_value = float(value)
                    current = numeric_series > numeric_value
                except:
                    current = pd.Series([False] * len(self.df))

            # ================= LESS THAN =================
            elif operator == "less than":
                numeric_series = pd.to_numeric(series, errors="coerce")

                try:
                    numeric_value = float(value)
                    current = numeric_series < numeric_value
                except:
                    current = pd.Series([False] * len(self.df))

            # ================= IN OPERATOR =================
            elif operator == "in":

                values_list = [
                    v.strip().lower()
                    for v in value.split(",")
                    if v.strip()
                ]

                current = series.str.lower().isin(values_list)

            # ================= DEFAULT =================
            else:
                current = pd.Series([False] * len(self.df))

            # IMPORTANT: fillna BEFORE combining
            condition &= current.fillna(False)

        matched_rows = self.df[condition]

        return matched_rows.copy()