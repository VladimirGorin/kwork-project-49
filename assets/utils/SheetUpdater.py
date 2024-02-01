from datetime import datetime
import gspread, time
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe

class SheetUpdater:
    def __init__(self, credentials_file, spreadsheet_id):
        # Устанавливаем необходимые права доступа и создаем объект авторизации
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        self.gc = gspread.authorize(credentials)

        # Открываем нужную таблицу по её идентификатору
        self.spreadsheet = self.gc.open_by_key(spreadsheet_id)

        # Устанавливаем задержку между запросами к API
        self.api_request_delay = 15  # Добавим задержку между запросами к API

    def update_sheet(self, closing_balance=None, local_df=None, start_date=None, end_date=None, desired_account_number=None, worksheet_index=None):
        worksheet = self.spreadsheet.get_worksheet(worksheet_index) if worksheet_index is not None else None

        # Обновление листа "Выписка"
        if worksheet_index == 0:
            try:
                set_with_dataframe(worksheet, local_df, 2, 15, include_column_header=False)
                print("Данные успешно установлены в таблицу Выписки.")
            except Exception as e:
                print(f"Ошибка при попытке установки данных в таблицу: {e}")

        # Обновление листа "Наличные"
        elif worksheet_index == 1:
            try:
                start_col = 6  # F
                end_col = worksheet.col_count  # Последняя колонка
                account_column_index = None

                # Находим индекс колонки с нужным счетом
                for col_index in range(start_col, end_col + 1):
                    account_numbers = worksheet.col_values(col_index)
                    if desired_account_number in account_numbers:
                        account_column_index = col_index
                        break

                if account_column_index:
                    cell_list = worksheet.range(2, account_column_index, worksheet.row_count, account_column_index)
                    start_date_cells = worksheet.col_values(3)[5:]
                    end_date_cells = worksheet.col_values(4)[5:]

                    for i, cell in enumerate(cell_list):
                        start_date_cell = datetime.strptime(start_date_cells[i], '%d.%m.%Y')
                        end_date_cell = datetime.strptime(end_date_cells[i], '%d.%m.%Y')

                        # Проверяем, попадают ли даты в интервал
                        if start_date <= start_date_cell <= end_date or start_date <= end_date_cell <= end_date:
                            worksheet.update_cell(i + 2, account_column_index, 'да')  # Обновляем ячейку
                            print(f"Успешно установлено значение 'да'\n")
                            break

                    print(f"Спим {self.api_request_delay} секунд что бы не попасть в лимит")
                    time.sleep(self.api_request_delay)  # Вводим задержку между запросами к API

            except Exception as e:
                print(f"Ошибка при попытке установки значение 'да': {e}")

        # Обновление листа "Остатки"
        elif worksheet_index == 2:
            try:
                start_col = 4
                end_col = worksheet.col_count
                account_column_index = None
                start_row = None
                end_row = None

                # Находим индекс колонки с нужным счетом
                for col_index in range(start_col, end_col + 1):
                    account_numbers = worksheet.col_values(col_index)
                    if desired_account_number in account_numbers:
                        account_column_index = col_index
                        break

                if account_column_index:
                    cell_list = worksheet.range(5, account_column_index, worksheet.row_count, account_column_index)
                    date_col = worksheet.col_values(3)[4:]

                    for i, cell in enumerate(cell_list):
                        current_date_cell = datetime.strptime(date_col[i], "%d.%m.%Y").date()
                        start_date_parsed = datetime.strptime(start_date.strftime("%d.%m.%Y"), "%d.%m.%Y").date()
                        end_date_parsed = datetime.strptime(end_date.strftime("%d.%m.%Y"), "%d.%m.%Y").date()

                        if start_date_parsed == current_date_cell:
                            start_row = i + 5
                        elif end_date_parsed == current_date_cell:
                            end_row = i + 5
                            break

                    if start_row is not None and end_row is not None:
                        for i in range(start_row - 5, end_row - 4):  # Корректируем индекс для соответствия cell_list
                            cell_list[i].value = closing_balance

                        worksheet.update_cells(cell_list)
                        print(f"Спим {self.api_request_delay} секунд что бы не попасть в лимит")
                        time.sleep(self.api_request_delay)  # Вводим задержку между запросами к API

            except Exception as e:
                print(f"Ошибка при обновлении листа: {e}")
