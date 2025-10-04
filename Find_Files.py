import os
import pathlib
import shutil
import time

import pandas as pd
import threading
import traceback

from natsort import natsorted
from PyQt5.QtCore import QThread, pyqtSignal
from DoingWindow import CheckWindow


class CancelException(Exception):
    pass


class FindingFiles(QThread):
    status_finish = pyqtSignal(str, str)
    progress_value = pyqtSignal(int)
    info_value = pyqtSignal(str, str)
    status = pyqtSignal(str)
    line_progress = pyqtSignal(str)
    line_doing = pyqtSignal(str)

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.unloading_file = incoming_data['unloading_file']
        self.start_path = incoming_data['start_path']
        self.finish_path = incoming_data['finish_path']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()
        self.event.set()
        self.now_doc = 0
        self.all_doc = 0
        self.percent_progress = 0
        self.move = incoming_data['move']
        self.name_dir = pathlib.Path(self.start_path).name
        title = f'Поиск файлов в папке «{self.name_dir}»'
        self.window_check = CheckWindow(self.default_path, self.event, self.move, title)
        self.progress_value.connect(self.window_check.progressBar.setValue)
        self.line_progress.connect(self.window_check.lineEdit_progress.setText)
        self.line_doing.connect(self.window_check.lineEdit_doing.setText)
        self.info_value.connect(self.window_check.info_message)
        self.window_check.show()

    def run(self):
        try:
            incoming_errors = []
            current_progress = 0
            self.logging.info('Начинаем заполнение документов, читаем значения из файла')
            files = [file for file in os.listdir(self.start_path)
                     if (file.endswith('.xlsx') or file.endswith('.docx')) and '~' not in file]
            files = natsorted(files)
            self.all_doc = len(files)
            self.percent_progress = 100 / len(files)
            self.line_progress.emit(f'Выполнено {int(current_progress)} %')
            self.progress_value.emit(0)
            if self.unloading_file.endswith('.txt'):
                df = pd.read_csv(self.unloading_file, delimiter='|', encoding='ANSI',
                                 header=None, converters={0: str, 6: str, 11: str})
                df.drop([i for i in range(df.shape[1]) if i not in [0, 6, 11]], axis=1, inplace=True)
                df = df[[0, 11, 6]]
                df.columns = range(df.shape[1])
            else:
                df = pd.read_excel(self.unloading_file, converters={'B': str, 'C': str}, header=None)
                df.columns = range(df.shape[1])
            self.logging.info('DataFrame заполнен, начинаем переименование и перенос')
            for file in files:
                self.now_doc += 1
                self.line_doing.emit(f'Генерируем файлы для комплекта {file} ({self.now_doc} из {self.all_doc})')
                self.event.wait()
                if self.window_check.stop_threading:
                    raise CancelException()
                self.logging.info(f'Поиск файла {file} ({self.now_doc} из {self.all_doc})')
                set_index = df[df[1] == file.rpartition('.')[0]].index.tolist()
                device_index = df[df[2] == file.rpartition('.')[0]].index.tolist()
                if not set_index and not device_index:
                    self.logging.info(f'{file} не найден в выгрузке, продолжаем')
                    current_progress += self.percent_progress
                    self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                    self.progress_value.emit(int(current_progress))
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
                    current_progress += self.percent_progress
                    self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                    self.progress_value.emit(int(current_progress))
                    continue
                try:
                    shutil.copy(pathlib.Path(self.start_path, file), pathlib.Path(self.finish_path, new_name))
                    self.logging.info(f'Скопировали файл в новую папку')
                except PermissionError:
                    self.logging.info(f'Не удалось скопировать файл')
                    incoming_errors.append(f'Не удалось скопировать файл {file}')
                    current_progress += self.percent_progress
                    self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                    self.progress_value.emit(int(current_progress))
                    continue
                try:
                    os.remove(pathlib.Path(self.start_path, file))
                    self.logging.info(f'Удалили старый файл')
                except PermissionError:
                    self.logging.info(f'Не удалось удалить файл')
                    incoming_errors.append(f'Не удалось удалить файл {file}')
                current_progress += self.percent_progress
                self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                self.progress_value.emit(int(current_progress))
            self.line_progress.emit(f'Выполнено 100 %')
            self.progress_value.emit(int(100))
            if incoming_errors:
                self.logging.info(f"Поиск файлов в папке «{self.name_dir}» завершён с ошибками")
                self.logging.info('\n'.join(incoming_errors))
                err = '\n' + '\n'.join(incoming_errors)
                self.info_value.emit('УПС!', f"Ошибки при работе программы:{err}")
                self.event.clear()
                self.event.wait()
                self.status.emit(f"Поиск файлов в папке «{self.name_dir}» завершён с ошибками")
                os.chdir(self.default_path)
                self.status_finish.emit('finding_files', str(self))
                time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
                self.window_check.close()
                return
            else:
                self.logging.info(f"Поиск файлов в папке «{self.name_dir}» успешно завершен")
                os.chdir(self.default_path)
                self.status.emit(f"Поиск файлов в папке «{self.name_dir}» успешно завершен")
                self.status_finish.emit('finding_files', str(self))
                time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
                self.window_check.close()
                # print(datetime.datetime.now() - start_time)
                return
        except CancelException:
            self.logging.warning(f"Поиск файлов в папке «{self.name_dir}» отменён пользователем")
            self.status.emit(f"Поиск файлов в папке «{self.name_dir}» отменён пользователем")
            os.chdir(self.default_path)
            self.status_finish.emit('finding_files', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.logging.warning(f"Поиск файлов в папке «{self.name_dir}» не завершён из-за ошибки")
            self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
            self.event.clear()
            self.event.wait()
            self.status.emit(f"Ошибка при поиске файлов в папке «{self.name_dir}»")
            os.chdir(self.default_path)
            self.status_finish.emit('finding_files', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
