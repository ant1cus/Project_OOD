# import datetime
import os
import pathlib
import random
import threading
import time

import traceback
import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal
from DoingWindow import CheckWindow


class CancelException(Exception):
    pass


class LFPEMIGeneration(QThread):
    status_finish = pyqtSignal(str, str)
    progress_value = pyqtSignal(int)
    info_value = pyqtSignal(str, str)
    status = pyqtSignal(str)
    line_progress = pyqtSignal(str)
    line_doing = pyqtSignal(str)

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.start_path = incoming_data['start_path']
        self.finish_path = incoming_data['finish_path']
        self.set_number = incoming_data['set_number']
        self.values_spread = incoming_data['values_spread']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()
        self.event.set()
        self.move = incoming_data['move']
        self.all_doc = 0
        self.now_doc = 0
        self.percent_progress = 0
        self.name_dir = pathlib.Path(self.start_path).name
        title = f'Генерация сплошного спектра в папке «{self.name_dir}»'
        self.window_check = CheckWindow(self.default_path, self.event, self.move, title)
        self.progress_value.connect(self.window_check.progressBar.setValue)
        self.line_progress.connect(self.window_check.lineEdit_progress.setText)
        self.line_doing.connect(self.window_check.lineEdit_doing.setText)
        self.info_value.connect(self.window_check.info_message)
        self.window_check.show()

    def run(self):
        self.progress_value.emit(0)
        try:
            self.progress_value.emit(100)
            self.status.emit(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» успешно завершена")
            self.logging.info(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» успешно завершена")
            os.chdir(self.default_path)
            self.status_finish.emit('gen_lf_pemi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            # print(datetime.datetime.now() - start_time)
            return
        except CancelException:
            self.logging.warning(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» отменена пользователем")
            self.status.emit(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» отменена пользователем")
            os.chdir(self.default_path)
            self.status_finish.emit('gen_lf_pemi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.logging.warning(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» не завершена из-за ошибки")
            self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
            self.event.clear()
            self.event.wait()
            self.status.emit(f"Ошибка при генерации сплошного спектра в папке «{self.name_dir}»")
            os.chdir(self.default_path)
            self.status_finish.emit('gen_lf_pemi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
