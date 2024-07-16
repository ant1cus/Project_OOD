# import datetime
import os
import pathlib
import threading

import shutil
import time
import traceback
from PyQt5.QtCore import QThread, pyqtSignal
from DoingWindow import CheckWindow


class CancelException(Exception):
    pass


class GenerateCopyApplication(QThread):
    status_finish = pyqtSignal(str, str)
    progress_value = pyqtSignal(int)
    info_value = pyqtSignal(str, str)
    status = pyqtSignal(str)
    line_progress = pyqtSignal(str)
    line_doing = pyqtSignal(str)

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.file = incoming_data['file']
        self.path = incoming_data['path']
        self.position_num = incoming_data['position_num']
        self.quantity = incoming_data['quantity']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()
        self.event.set()
        self.all_doc = 0
        self.now_doc = 0
        self.percent_progress = 0
        self.move = incoming_data['move']
        self.name_dir = pathlib.Path(self.path).name
        title = f'Создание копий приложений в папке «{self.name_dir}»'
        self.window_check = CheckWindow(self.default_path, self.event, self.move, title)
        self.progress_value.connect(self.window_check.progressBar.setValue)
        self.line_progress.connect(self.window_check.lineEdit_progress.setText)
        self.line_doing.connect(self.window_check.lineEdit_doing.setText)
        self.info_value.connect(self.window_check.info_message)
        self.window_check.show()

    def run(self):
        try:
            name = os.path.basename(self.file)
            name_suff = name.rpartition('.')[0].title()
            current_progress = 0
            self.logging.info('Начинаем копировать документы')
            self.line_progress.emit(f'Выполнено {int(current_progress)} %')
            self.progress_value.emit(0)
            self.percent_progress = 100 / self.quantity
            self.all_doc = self.quantity
            for number in range(1, self.quantity + 1):
                self.now_doc += 1
                self.line_doing.emit(f'Создаем документ {self.position_num}.{self.position_num}'
                                     f' ({self.now_doc} из {self.all_doc})')
                self.event.wait()
                if self.window_check.stop_threading:
                    raise CancelException()
                self.logging.info(f'Создаем документ {self.position_num}.{self.position_num}'
                                  f' ({self.now_doc} из {self.all_doc})')
                shutil.copy(self.file, self.path)
                os.rename(self.path + '\\' + name, self.path + '\\' + name_suff + ' ' + str(self.position_num) +
                          '.' + str(number) + '.docx')
                current_progress += self.percent_progress
                self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                self.progress_value.emit(int(current_progress))
            self.line_progress.emit(f'Выполнено 100 %')
            self.logging.info(f"Создание копий приложений в папке «{self.name_dir}» успешно завершено")
            self.progress_value.emit(int(100))
            os.chdir(self.default_path)
            self.status.emit(f"Создание копий приложений в папке «{self.name_dir}» успешно завершено")
            self.status_finish.emit('copy_application', str(self))
            time.sleep(0.1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            # print(datetime.datetime.now() - start_time)
            return
        except CancelException:
            self.logging.warning(f"Создание копий приложений в папке «{self.name_dir}» отменено пользователем")
            self.status.emit(f"Создание копий приложений в папке «{self.name_dir}» отменено пользователем")
            os.chdir(self.default_path)
            self.status_finish.emit('copy_application', str(self))
            time.sleep(0.1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.logging.warning(f"Создание копий приложений в папке «{self.name_dir}» не заврешено из-за ошибки")
            self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
            self.event.clear()
            self.event.wait()
            self.status.emit(f"Ошибка при создании копий приложений в папке «{self.name_dir}»")
            os.chdir(self.default_path)
            self.status_finish.emit('copy_application', str(self))
            time.sleep(0.1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
