import random
import os
import threading
import traceback

from PyQt5.QtCore import QThread, pyqtSignal


class HFIGeneration(QThread):
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
        self.mode = incoming_data['mode']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()

    def run(self):
        def create(name, frq, rl, path_for_file):
            with open(path_for_file + '\\' + name + '.txt', mode='w') as f:
                print('{0:<} {1:>} {2:>}'.format(frq, '1000000', '0'), file=f)
                for j in rl:
                    gen = "%.4f" % random.uniform(0.071, 0.121)
                    print('{0:<} {1:>} {2:>} {3:>}'.format('-', j, gen, '0'), file=f)

        try:
            current_progress = 0
            self.logging.info('Начинаем генерацию ВЧН')
            self.status.emit('Старт')
            self.progress.emit(current_progress)
            percent = 100/self.quantity
            mode_name = ['Питание', 'Симметричка', 'Несимметричка']
            ref_lvl = [200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000,
                       10000]
            for i in range(1, int(self.quant) + 1, 1):
                self.logging.info('Генерируем файлы для комплекта ' + str(i))
                self.status.emit('Генерируем файлы для комплекта ' + str(i))
                os.makedirs(self.path + '\\' + str(i))
                if self.mode[0]:
                    create(mode_name[0], self.freq, ref_lvl, self.path + '\\' + str(i))
                if self.mode[1]:
                    create(mode_name[1], self.freq, ref_lvl, self.path + '\\' + str(i))
                if self.mode[2]:
                    create(mode_name[2], self.freq, ref_lvl, self.path + '\\' + str(i))
                current_progress += percent
                self.progress.emit(int(current_progress))
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

