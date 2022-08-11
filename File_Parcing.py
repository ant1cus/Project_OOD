import os
import threading
import traceback
from PyQt5.QtCore import QThread, pyqtSignal
from convert import file_parcing


class FileParcing(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.path = incoming_data['path']
        self.all_progress = incoming_data['progress']
        self.group_file = incoming_data['group_file']
        self.logging = incoming_data['logging']
        self.q = incoming_data['q']
        self.event = threading.Event()

    def run(self):
        try:
            current_progress = 0
            self.logging.info("Начинаем")
            self.status.emit('Старт')
            self.progress.emit(current_progress)
            percent = 100/self.all_progress
            errors = []
            succsess_path = []
            error_path = []
            if self.group_file:
                for folder in os.listdir(self.path):
                    if os.path.isdir(self.path + '\\' + folder):
                        err = file_parcing(self.path + '\\' + folder, self.logging, self.status, self.progress, percent,
                                           current_progress)
                        if err['error']:
                            error_path.append(self.path + '\\' + folder)
                            for element in err['error']:
                                errors.append(element)
                        else:
                            succsess_path.append(self.path + '\\' + folder)
                        current_progress = err['cp']
            else:
                err = file_parcing(self.path, self.logging, self.status, self.progress, percent, current_progress)
                if err['error']:
                    errors = err['error']
                    error_path.append(self.path)
                else:
                    succsess_path.append(self.path)
            if succsess_path:
                self.q.put({'Прошедшие заказы:': '\n'.join(succsess_path)})
                self.errors.emit()
            if error_path:
                self.q.put({'Заказы с ошибками:': '\n'.join(error_path)})
                self.errors.emit()
            if errors:
                self.logging.info("Выводим ошибки")
                self.q.put({'errors': errors})
                self.errors.emit()
                self.status.emit('В файлах присутствуют ошибки')
            else:
                self.logging.info("Конец работы программы")
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

