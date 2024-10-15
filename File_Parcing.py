import os
import pathlib
import threading
import time
import traceback
from PyQt5.QtCore import QThread, pyqtSignal
from convert import file_parcing
from DoingWindow import CheckWindow


class CancelException(Exception):
    pass


class FileParcing(QThread):
    status_finish = pyqtSignal(str, str)
    progress_value = pyqtSignal(int)
    info_value = pyqtSignal(str, str)
    status = pyqtSignal(str)
    line_progress = pyqtSignal(str)
    line_doing = pyqtSignal(str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.path = incoming_data['path']
        self.all_progress = incoming_data['progress']
        self.group_file = incoming_data['group_file']
        self.no_freq_lim = incoming_data['no_freq_lim']
        self.twelve_sectors = incoming_data['12_sec']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()
        self.event.set()
        self.now_doc = 0
        self.all_doc = 0
        self.percent_progress = 0
        self.move = incoming_data['move']
        self.name_dir = pathlib.Path(self.path).name
        title = f'Парсинг файлов в папке «{self.name_dir}»'
        self.window_check = CheckWindow(self.default_path, self.event, self.move, title)
        self.progress_value.connect(self.window_check.progressBar.setValue)
        self.line_progress.connect(self.window_check.lineEdit_progress.setText)
        self.line_doing.connect(self.window_check.lineEdit_doing.setText)
        self.info_value.connect(self.window_check.info_message)
        self.window_check.show()

    def run(self):
        try:
            current_progress = 0
            self.logging.info('Начинаем парсить файлы')
            self.all_doc = self.all_progress
            self.percent_progress = 100 / self.all_progress
            self.line_progress.emit(f'Выполнено {int(current_progress)} %')
            self.progress_value.emit(0)
            errors = []
            succsess_path = []
            error_path = []
            if self.group_file:
                for folder in os.listdir(self.path):
                    self.event.wait()
                    if self.window_check.stop_threading:
                        raise CancelException()
                    if os.path.isdir(pathlib.Path(self.path, folder)):
                        err = file_parcing(pathlib.Path(self.path, folder), self.logging, self.line_doing, self.now_doc,
                                           self.all_doc, self.line_progress, self.progress_value, self.percent_progress,
                                           current_progress, self.no_freq_lim, self.default_path, self.event,
                                           self.window_check, self.twelve_sectors)
                        if err['base_exception']:
                            self.logging.error(err['text'])
                            self.logging.error(err['trace'])
                            self.logging.warning(f"Парсинг файлов в папке «{self.name_dir}» не завершён из-за ошибки")
                            self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
                            self.event.clear()
                            self.event.wait()
                            self.status.emit(f"Ошибка при парсинге файлов в папке «{self.name_dir}»")
                            os.chdir(self.default_path)
                            self.status_finish.emit('parcing_file', str(self))
                            time.sleep(0.1)  # Не удалять, не успевает отработать emit status_finish. Может потом
                            self.window_check.close()
                            return
                        if err['cancel']:
                            raise CancelException()
                        if err['error']:
                            error_path.append(pathlib.Path(self.path, folder))
                            for element in err['error']:
                                errors.append(element)
                        else:
                            succsess_path.append(pathlib.Path(self.path, folder))
                        current_progress = err['cp']
                        self.now_doc = err['now_doc']
            else:
                err = file_parcing(self.path, self.logging, self.line_doing, self.now_doc, self.all_doc,
                                   self.line_progress, self.progress_value, self.percent_progress, current_progress,
                                   self.no_freq_lim, self.default_path, self.event, self.window_check,
                                   self.twelve_sectors)
                if err['base_exception']:
                    self.logging.error(err['text'])
                    self.logging.error(err['trace'])
                    self.logging.warning(f"Парсинг файлов в папке «{self.name_dir}» не завершён из-за ошибки")
                    self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
                    self.event.clear()
                    self.event.wait()
                    self.status.emit(f"Ошибка при парсинге файлов в папке «{self.name_dir}»")
                    os.chdir(self.default_path)
                    self.status_finish.emit('parcing_file', str(self))
                    time.sleep(0.1)  # Не удалять, не успевает отработать emit status_finish. Может потом
                    self.window_check.close()
                    return
                if err['cancel']:
                    raise CancelException
                if err['error']:
                    errors = err['error']
                    error_path.append(self.path)
                else:
                    succsess_path.append(self.path)
            if succsess_path:
                self.queue.put({'Прошедшие заказы:': '\n'.join(succsess_path)})
                self.errors.emit()
            if error_path:
                self.queue.put({'Заказы с ошибками:': '\n'.join(error_path)})
                self.errors.emit()
            if errors:
                self.logging.info("Выводим ошибки")
                self.queue.put({'errors': errors})
                self.errors.emit()
                self.status.emit('В файлах присутствуют ошибки')
            else:
                self.logging.info("Конец работы программы")
                self.status.emit('Готово')
            self.line_progress.emit(f'Выполнено 100 %')
            self.progress_value.emit(int(100))
            self.logging.info(f"Генрация файлов в папке «{self.name_dir}» успешно завершена")
            os.chdir(self.default_path)
            self.status.emit(f"Парсинг файлов в папке «{self.name_dir}» успешно завершён")
            self.status_finish.emit('parcing_file', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            # print(datetime.datetime.now() - start_time)
            return
        except CancelException:
            self.logging.warning(f"Парсинг файлов в папке «{self.name_dir}» отменён пользователем")
            self.status.emit(f"Парсинг файлов в папке «{self.name_dir}» отменён пользователем")
            os.chdir(self.default_path)
            self.status_finish.emit('parcing_file', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.logging.warning(f"Парсинг файлов в папке «{self.name_dir}» не завершён из-за ошибки")
            self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
            self.event.clear()
            self.event.wait()
            self.status.emit(f"Ошибка при парсинге файлов в папке «{self.name_dir}»")
            os.chdir(self.default_path)
            self.status_finish.emit('parcing_file', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return

