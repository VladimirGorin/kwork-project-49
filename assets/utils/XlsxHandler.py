import pandas as pd
import csv
import warnings


class XlsxHandler:
    def __init__(self, xlsx_file_path, output_csv_path):
        self.xlsx_file_path = xlsx_file_path
        self.output_csv_path = output_csv_path
        self.sheet_names = []
        self.local_df = None

    def extract_xlsx_to_csv(self, sheet_id=0):
        try:
            na_values = ['', None, 'NaN', 'N/A', 'NA', 'na', 'n/a']

            with warnings.catch_warnings(record=True):
                warnings.simplefilter("always")
                xls = pd.ExcelFile(self.xlsx_file_path)
                self.sheet_names = xls.sheet_names

                df = pd.read_excel(self.xlsx_file_path, engine="openpyxl",
                                   header=None, na_values=na_values, sheet_name=sheet_id)
                self.local_df = df

            df = df.map(lambda x: '' if pd.isna(x) and x != '' else x)
            df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
            df = df.dropna(axis=1, how='all')

            df.to_csv(self.output_csv_path, index=False,
                      header=False, na_rep='')

            print(
                f"Данные успешно сохранены в CSV-файл: {self.output_csv_path}")

        except Exception as e:
            print(f"Ошибка при чтении файла XLSX или сохранении в CSV: {e}")

    def get_csv_data(self):
        try:
            with open(self.output_csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                result_array = []

                csv_reader = csv.reader(csvfile)

                non_empty_rows = filter(lambda row: any(
                    cell.strip() for cell in row), csv_reader)

                for row in non_empty_rows:
                    filter_row = [cell for cell in row if cell.strip()]

                    result_array.append(filter_row)

            return result_array

        except Exception as e:
            print(f"Ошибка при чтении файла CSV: {e}")
