import os
import shutil
import threading
import traceback

from PyQt5.QtCore import QThread, pyqtSignal


class GenerateCopyApplication(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.file = incoming_data['file']
        self.path = incoming_data['path']
        self.position_num = incoming_data['position_num']
        self.quantity = incoming_data['quantity']
        self.logging = incoming_data['logging']
        self.q = incoming_data['q']
        self.event = threading.Event()

    def run(self):
        try:
            name = os.path.basename(self.file)
            name_suff = name.rpartition('.')[0].title()
            current_progress = 0
            self.logging.info('Начинаем копировать документы')
            self.status.emit('Старт')
            self.progress.emit(current_progress)
            percent = 100 / self.quantity
            for number in range(1, self.quantity + 1):
                self.logging.info('Создаем документ ' + str(self.position_num) + '.' + str(number))
                self.status.emit('Создаем документ ' + str(self.position_num) + '.' + str(number))
                shutil.copy(self.file, self.path)
                os.rename(self.path + '\\' + name, self.path + '\\' + name_suff + ' ' + str(self.position_num) +
                          '.' + str(number) + '.docx')
                current_progress += percent
                self.progress.emit(current_progress)
            self.logging.info("Конец работы программы")
            self.progress.emit(100)
            self.status.emit('Готово')
            os.chdir('C:\\')
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.progress.emit(0)
            self.status.emit('Ошибка!')
            os.chdir('C:\\')
            return
