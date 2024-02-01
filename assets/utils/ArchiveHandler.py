import os, zipfile, shutil


class ArchiveHandler:
    def __init__(self):
        self.archive_path = './assets/data/archives'
        self.output_path = './assets/data/output'
        self.log_file_path = './assets/data/processed_archives.txt'

    def extract_xlsx_files(self):
        if not os.path.exists(self.archive_path):
            print(f"Папка с архивами '{self.archive_path}' не существует.")
            return

        if not os.path.exists(self.output_path):
            os.makedirs(self.output_path)

        for file_name in os.listdir(self.archive_path):
            file_path = os.path.join(self.archive_path, file_name)

            if file_name.endswith('.zip') and not self.is_archive_processed(file_name):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    for member in zip_ref.infolist():
                        if member.filename.endswith('.xlsx'):
                            output_file_path = os.path.join(self.output_path, os.path.basename(member.filename))
                            with zip_ref.open(member) as source, open(output_file_path, 'wb') as target:
                                shutil.copyfileobj(source, target)
                            print(f"Извлечен файл: {member.filename} в {output_file_path}")

                self.log_processed_archive(file_name)

    def is_archive_processed(self, archive_name):
        if os.path.exists(self.log_file_path):
            with open(self.log_file_path, 'r') as log_file:
                processed_archives = log_file.read().splitlines()
                return archive_name in processed_archives
        return False

    def log_processed_archive(self, archive_name):
        with open(self.log_file_path, 'a') as log_file:
            log_file.write(archive_name + '\n')

    def get_all_xlsx_paths(self):
        xlsx_paths = []
        if os.path.exists(self.output_path):
            for file_name in os.listdir(self.output_path):
                file_path = os.path.join(self.output_path, file_name)
                if os.path.isfile(file_path) and file_name.endswith('.xlsx'):
                    xlsx_paths.append(file_path)
        return xlsx_paths

    def delete_all_files_in_output(self):
        output_path = os.path.abspath(self.output_path)

        # Проверка существования папки
        if not os.path.exists(output_path):
            print(f"Папка {self.output_path} не существует.")
            return

        # Получение списка файлов в папке
        files = os.listdir(output_path)

        # Удаление каждого файла
        for file_name in files:
            file_path = os.path.join(output_path, file_name)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    print(f"Файл {file_name} успешно удален.")
                else:
                    print(f"{file_name} не является файлом.")
            except Exception as e:
                print(f"Ошибка при удалении файла {file_name}: {e}")
