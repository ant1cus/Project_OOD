import os
import pathlib
import shutil

import pandas as pd
import threading
import traceback

from natsort import natsorted
from PyQt5.QtCore import QThread, pyqtSignal


class FindingFiles(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.unloading_file = incoming_data['unloading_file']
        self.start_path = incoming_data['start_path']
        self.finish_path = incoming_data['finish_path']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()

    def run(self):
        try:
            incoming_errors = []
            progress = 0
            self.logging.info('Начинаем заполнение документов')
            self.status.emit('Считываем значения из файла')
            self.progress.emit(progress)
            self.logging.info('Получаем список файлов в папке')
            files = [file for file in os.listdir(self.start_path) if file.endswith('.xlsx') and '~' not in file]
            files = natsorted(files)
            percent = 100 / len(files)
            if self.unloading_file.endswith('.txt'):
                df = pd.read_csv(self.unloading_file, delimiter='|', encoding='ANSI',
                                 header=None, converters={6: str, 11: str})
                df.drop([i for i in range(df.shape[1]) if i not in [0, 6, 11]], axis=1, inplace=True)
                df = df[[0, 11, 6]]
                df.columns = range(df.shape[1])
            else:
                df = pd.read_excel(self.unloading_file, converters={'B': str, 'C': str}, header=None)
                df.columns = range(df.shape[1])
            self.logging.info('DataFrame заполнен, начинаем переименование и перенос')
            for file in files:
                self.logging.info(f'Поиск файла {file}')
                self.status.emit(f'Обработка файла {file}')
                set_index = df[df[1] == file.rpartition('.')[0]].index.tolist()
                device_index = df[df[2] == file.rpartition('.')[0]].index.tolist()
                if not set_index and not device_index:
                    self.logging.info(f'{file} не найден в выгрузке, продолжаем')
                    progress += percent
                    self.progress.emit(int(progress))
                    continue
                if len(set_index) > 0:
                    self.logging.info(f'Берем имя в номере комплекта')
                    new_name = str(df.iloc[set_index[0], 0])
                else:
                    self.logging.info(f'Берем имя в номере устройства')
                    new_name = str(df.iloc[device_index[0], 0])
                new_name = new_name.partition('.')[2] + '.xlsx'
                if new_name is False:
                    self.logging.info(f'Не удалось извлечь новое имя файла')
                    incoming_errors.append(f'Неверно указан номер комплекта у файла {file}'
                                           f'(отсутствует нужный разделитель «.»)')
                    progress += percent
                    self.progress.emit(int(progress))
                    continue
                try:
                    shutil.copy(pathlib.Path(self.start_path, file), pathlib.Path(self.finish_path, new_name))
                    self.logging.info(f'Скопировали файл в новую папку')
                except PermissionError:
                    self.logging.info(f'Не удалось скопировать файл')
                    incoming_errors.append(f'Не удалось скопировать файл {file}')
                    progress += percent
                    self.progress.emit(int(progress))
                    continue
                try:
                    os.remove(pathlib.Path(self.start_path, file))
                    self.logging.info(f'Удалили старый файл')
                except PermissionError:
                    self.logging.info(f'Не удалось удалить файл')
                    incoming_errors.append(f'Не удалось удалить файл {file}')
                progress += percent
                self.progress.emit(int(progress))
            self.logging.info("Конец работы программы")
            self.progress.emit(100)
            self.status.emit('Готово')
            os.chdir(self.default_path)
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.progress.emit(0)
            self.status.emit('Ошибка!')
            os.chdir(self.default_path)
            return
