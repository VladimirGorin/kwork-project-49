from assets.utils.ArchiveHandler import ArchiveHandler
from assets.utils.XlsxHandler import XlsxHandler
from assets.utils.SheetUpdater import SheetUpdater
from datetime import datetime

import configparser


def print_separator():
    print("\n" + "=" * 50 + "\n")


# Статистика
stats = {
    "gived_files": 0,
    "processed data": 0
}

# Читаем конфигурацию
config = configparser.ConfigParser()
archive_handler = ArchiveHandler()
config.read('config.ini')

# Загружаем параметры конфигурации
CREDENTIALS_FILE = config.get('Sheet', 'CREDENTIALS_FILE')
SPREADSHEET_ID = config.get('Sheet', 'SPREADSHEET_ID')
csv_file_path = './assets/data/output/converted_data.csv'

# Инициализируем SheetUpdater
updater = SheetUpdater(credentials_file=CREDENTIALS_FILE,
                       spreadsheet_id=SPREADSHEET_ID)

# Инициализируем ArchiveHandler и извлекаем XLSX-файлы
archive_handler.extract_xlsx_files()
all_xlsx_paths = archive_handler.get_all_xlsx_paths()

stats["gived_files"] = len(all_xlsx_paths)

# Итерируемся по XLSX-файлам
for xlsx_file_path in all_xlsx_paths:
    print_separator()
    print(f"Работа с файлом: {xlsx_file_path}")

    # Извлекаем XLSX в CSV с использованием XlsxHandler
    xlsx_handler = XlsxHandler(xlsx_file_path, csv_file_path)
    xlsx_handler.extract_xlsx_to_csv()

    print(xlsx_handler.sheet_names)

    # Проверяем, является ли первый лист выпиской
    if "Выписка" in xlsx_handler.sheet_names[0]:
        print_separator()
        print(f"Обработка листа 'Выписка': {xlsx_handler.sheet_names[0]}")
        # Извлекаем данные CSV
        csv_data = xlsx_handler.get_csv_data()

        # Извлекаем и выводим дату
        try:
            date_str = csv_data[0][0]
            print(f"Исходная строка даты: {date_str}")

            cut_date_str = date_str.replace(
                "Выписка по счету с", "").replace("по", "").split()
            start_date = datetime.strptime(cut_date_str[0], '%d.%m.%Y')
            end_date = datetime.strptime(cut_date_str[1], '%d.%m.%Y')

            print(f"Обработанные даты: {start_date}, {end_date}")

        except Exception as e:
            print(f"Ошибка при разборе даты: {e}")
            continue

        # Извлекаем и выводим номер счета
        try:
            desired_account_number_str = csv_data[4][0]
            desired_cut_account_number_str = desired_account_number_str.replace(
                "Счет:", "").replace("(РУБ)", "").split()[0]

            if desired_cut_account_number_str.isdigit():
                print(f"Номер счета: {desired_cut_account_number_str}")
            else:
                print(f"Ошибка при извлечении номера счета. Содержит буквы: {
                      desired_cut_account_number_str}")
                continue

        except Exception as e:
            print(f"Ошибка при извлечении номера счета: {e}")
            continue

        # Извлекаем и выводим исходящий остаток
        try:
            closing_balance_str = csv_data[8][1].split(".")[0]
            print(f"Исходящий остаток: {closing_balance_str}")

            # Раскомментируйте следующую строку для обновления листа
            updater.update_sheet(start_date=start_date, closing_balance=closing_balance_str,
                                 end_date=end_date, desired_account_number=desired_cut_account_number_str, worksheet_index=2)
            updater.update_sheet(start_date=start_date, end_date=end_date,
                                 desired_account_number=desired_cut_account_number_str, worksheet_index=1)

        except Exception as e:
            print(f"Ошибка при извлечении исходящего остатка: {e}")
            continue

        print_separator()

    # Проверяем, является ли второй лист выпиской
    try:
        if "Выписка" in xlsx_handler.sheet_names[1]:
            print_separator()
            print(f"Обработка листа 'Выписка': {xlsx_handler.sheet_names[1]}")

            # Извлекаем данные CSV
            xlsx_handler.extract_xlsx_to_csv(sheet_id=1)
            csv_data = xlsx_handler.get_csv_data()

            # Извлекаем и выводим дату
            try:
                date_str = csv_data[0][0]
                cut_date_str = date_str.replace(
                    "Выписка по счету с", "").replace("по", "").split()

                start_date = datetime.strptime(cut_date_str[0], '%d.%m.%Y')
                end_date = datetime.strptime(cut_date_str[1], '%d.%m.%Y')

                print(f"Обработанные даты: {start_date}, {end_date}")

            except Exception as e:
                print(f"Ошибка при разборе даты: {e}")
                continue

            # Извлекаем и выводим номер счета
            try:
                desired_account_number_str = csv_data[3][0]
                desired_cut_account_number_str = desired_account_number_str.replace(
                    "Счет:", "").replace("(РУБ)", "").split()[0]

                if desired_cut_account_number_str.isdigit():
                    print(f"Номер счета: {desired_cut_account_number_str}")
                else:
                    print(f"Ошибка при извлечении номера счета. Содержит буквы: {
                          desired_cut_account_number_str}")
                    continue

            except Exception as e:
                print(f"Ошибка при извлечении номера счета: {e}")
                continue

            # Извлекаем и выводим исходящий остаток
            try:
                closing_balance_str = csv_data[7][1].split(".")[0]
                print(f"Исходящий остаток: {closing_balance_str}")

                # Раскомментируйте следующие строки для обновления листа
                updater.update_sheet(start_date=start_date, closing_balance=closing_balance_str,
                                     end_date=end_date, desired_account_number=desired_cut_account_number_str, worksheet_index=2)
                updater.update_sheet(start_date=start_date, end_date=end_date,
                                     desired_account_number=desired_cut_account_number_str, worksheet_index=1)

            except Exception as e:
                print(f"Ошибка при извлечении исходящего остатка: {e}")
                continue

            print_separator()
    except IndexError:
        ""

    # Проверяем, содержится ли "Операции" на обоих листах
    try:
        if "Операции" in xlsx_handler.sheet_names[0] or "Операции" in xlsx_handler.sheet_names[1]:
            xlsx_handler.extract_xlsx_to_csv()

            try:
                local_df = xlsx_handler.local_df.iloc[2:]
                print_separator()
                print("Данные успешно скопированы из таблицы 'Операции'")

            except Exception as e:
                print(f"Ошибка при чтении листа 'Операции': {e}")
                continue

            try:
                print("Поиск и обновление значений.\n")
                updater.update_sheet(local_df=local_df, worksheet_index=0)
            except Exception as e:
                print(f"Ошибка при записи в таблицу 'Выписка': {e}")
                continue
    except IndexError:
        ""

    stats["processed data"] += 1

archive_handler.delete_all_files_in_output()
print("Выполнение кода завершено.")
print_separator()
print(f"Статистика:\nВсего получено файлов: {
      stats['gived_files']}\nВсего обработано {stats['processed data']}")
print_separator()
