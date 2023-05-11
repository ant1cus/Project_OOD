import os
import pathlib
import random
import threading
import traceback

import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal
from convert import file_parcing


class GenerationFileCC(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.source = incoming_data['source']
        self.output = incoming_data['output']
        self.complect = incoming_data['complect']
        self.complect_quant = incoming_data['complect_quant']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()

    def run(self):
        try:
            current_progress = 0
            errors = []
            self.logging.info('Считывание файлов')
            self.status.emit('Старт')
            self.progress.emit(current_progress)
            for folder in os.listdir(self.source):
                for file_csv in os.listdir(str(pathlib.Path(self.source, folder))):
                    df = pd.read_csv(str(pathlib.Path(self.source, folder, file_csv)), delimiter=None,
                                     encoding="unicode_escape")
            if errors:
                self.logging.info("Выводим ошибки")
                self.queue.put({'errors_gen': errors})
                self.errors.emit()
                self.status.emit('В файлах присутствуют ошибки')
                self.progress.emit(0)
            else:
                self.progress.emit(100)
                self.logging.info("Конец работы программы")
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

