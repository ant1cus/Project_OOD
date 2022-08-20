import random
import os
import threading
import traceback

from PyQt5.QtCore import QThread, pyqtSignal


class HFEGeneration(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.path = incoming_data['path']
        self.quantity = incoming_data['quantity']
        self.freq = incoming_data['freq']
        self.val = incoming_data['val']
        self.logging = incoming_data['logging']
        self.q = incoming_data['q']
        self.event = threading.Event()

    def run(self):
        try:
            current_progress = 0
            self.logging.info('Начинаем генерацию ВЧО')
            self.status.emit('Старт')
            self.progress.emit(current_progress)
            percent = 100/self.quantity
            ref_lvl = [200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000,
                       10000]
            for i in range(1, self.quantity + 1, 1):
                self.logging.info('Генерируем файлы для комплекта ' + str(i))
                self.status.emit('Генерируем файлы для комплекта ' + str(i))
                os.chdir(self.path)
                os.makedirs(self.path + '\\' + str(i))
                with open(self.path + '\\' + str(i) + '\\' + 'ВЧО' + '.txt', mode='w') as f:
                    print('{0:<} {1:>} {2:>} {3:>} {4:>}'.format(self.freq, self.val, '1', '78', '1'), file=f)
                    for j in ref_lvl:
                        generation = random.uniform(0.021, 0.053)
                        gen = "%.4f" % float(generation)
                        print('{0:<} {1:>} {2:>} {3:>}'.format('-', j, gen, '0'), file=f)
                current_progress += percent
                self.pbar.setValue(current_progress)
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

