import pathlib
import random
import os
import threading
import time
import traceback

from PyQt5.QtCore import QThread, pyqtSignal
from DoingWindow import CheckWindow


class CancelException(Exception):
    pass


class HFIGeneration(QThread):
    status_finish = pyqtSignal(str, str)
    progress_value = pyqtSignal(int)
    info_value = pyqtSignal(str, str)
    status = pyqtSignal(str)
    line_progress = pyqtSignal(str)
    line_doing = pyqtSignal(str)

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
        self.event.set()
        self.now_doc = 0
        self.all_doc = 0
        self.percent_progress = 0
        self.move = incoming_data['move']
        self.name_dir = pathlib.Path(self.path).name
        title = f'Генерация ВЧН в папке «{self.name_dir}»'
        self.window_check = CheckWindow(self.default_path, self.event, self.move, title)
        self.progress_value.connect(self.window_check.progressBar.setValue)
        self.line_progress.connect(self.window_check.lineEdit_progress.setText)
        self.line_doing.connect(self.window_check.lineEdit_doing.setText)
        self.info_value.connect(self.window_check.info_message)
        self.window_check.show()

    def run(self):
        def create(name, frq, rl, path_for_file):
            with open(pathlib.Path(path_for_file, name + '.txt'), mode='w') as f:
                print('{0:<} {1:>} {2:>}'.format(frq, '1000000', '0'), file=f)
                for j in rl:
                    gen = "%.4f" % random.uniform(0.071, 0.121)
                    print('{0:<} {1:>} {2:>} {3:>}'.format('-', j, gen, '0'), file=f)

        try:
            current_progress = 0
            self.logging.info('Начинаем генерацию ВЧН')
            self.all_doc = self.quantity
            self.line_progress.emit(f'Выполнено {int(current_progress)} %')
            self.progress_value.emit(0)
            self.percent_progress = 100/self.quantity
            mode_name = ['Питание', 'Симметричка', 'Несимметричка']
            ref_lvl = [200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000,
                       10000]
            for i in range(1, int(self.quantity) + 1, 1):
                self.now_doc += 1
                self.line_doing.emit(f'Генерируем файлы для комплекта {i} ({self.now_doc} из {self.all_doc})')
                self.event.wait()
                if self.window_check.stop_threading:
                    raise CancelException()
                self.logging.info(f'Генерируем файлы для комплекта {i} ({self.now_doc} из {self.all_doc})')
                os.makedirs(self.path + '\\' + str(i))
                if self.mode[0]:
                    create(mode_name[0], self.freq, ref_lvl, self.path + '\\' + str(i))
                if self.mode[1]:
                    create(mode_name[1], self.freq, ref_lvl, self.path + '\\' + str(i))
                if self.mode[2]:
                    create(mode_name[2], self.freq, ref_lvl, self.path + '\\' + str(i))
                current_progress += self.percent_progress
                self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                self.progress_value.emit(int(current_progress))
            self.line_progress.emit(f'Выполнено 100 %')
            self.logging.info(f"Генерация ВЧН в папке «{self.name_dir}» успешно завершена")
            self.progress_value.emit(int(100))
            os.chdir(self.default_path)
            self.status.emit(f"Генерация ВЧН в папке «{self.name_dir}» успешно завершена")
            self.status_finish.emit('generate_hfi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            # print(datetime.datetime.now() - start_time)
            return
        except CancelException:
            self.logging.warning(f"Генерация ВЧН в папке «{self.name_dir}» отменена пользователем")
            self.status.emit(f"Генерация ВЧН в папке «{self.name_dir}» отменена пользователем")
            os.chdir(self.default_path)
            self.status_finish.emit('generate_hfi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.logging.warning(f"Генерация ВЧН в папке «{self.name_dir}» не заврешена из-за ошибки")
            self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
            self.event.clear()
            self.event.wait()
            self.status.emit(f"Ошибка при генерации ВЧН в папке «{self.name_dir}»")
            os.chdir(self.default_path)
            self.status_finish.emit('generate_hfi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return

