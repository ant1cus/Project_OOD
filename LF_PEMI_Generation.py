# import datetime
import os
import random
import re
import threading
import time

import traceback
import pandas as pd
import numpy as np
from pathlib import Path
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
        self.values_spread = float(incoming_data['values_spread'])
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()
        self.event.set()
        self.move = incoming_data['move']
        self.all_doc = len(self.set_number)
        self.now_doc = 0
        self.percent_progress = 100/self.all_doc
        self.name_dir = Path(self.start_path).name
        title = f'Генерация сплошного спектра в папке «{self.name_dir}»'
        self.window_check = CheckWindow(self.default_path, self.event, self.move, title)
        self.progress_value.connect(self.window_check.progressBar.setValue)
        self.line_progress.connect(self.window_check.lineEdit_progress.setText)
        self.line_doing.connect(self.window_check.lineEdit_doing.setText)
        self.info_value.connect(self.window_check.info_message)
        self.window_check.show()

    def run(self):
        try:
            max_difference = 3  # Для подстановки требуемой разницы
            self.progress_value.emit(0)
            current_progress = 0
            lines_file = ['n', 'p', 'z']
            frames = {i: pd.DataFrame(columns=['frq', 'max_s', 'min_s', 'max_n', 'min_n'])
                      for i in ['void', 'f', 'g', 'i', 'j', 'm', 'n', 'p', 'z']}
            errors = []
            for file in [f for f in Path(self.start_path).rglob("*.txt") if 'динамики' in f.name.lower()]:
                file_type = re.findall(r'\w+.(\w)\.txt', file.name)[0]
                if file_type == 'и':
                    file_type = 'void'
                df = pd.read_csv(file, sep='\t', header=None)
                if file_type in lines_file:
                    frq = df[0].to_numpy().tolist()
                    diff = (df[1].to_numpy() - df[2].to_numpy()).tolist()
                    if any(i >= max_difference for i in diff):
                        diff_larger_3 = [str(f) for f, d in zip(frq, diff) if d > 3]
                        # Надо найти частоту на которой не выполняется условие
                        errors.append(f"В файле {file.name} ({file.parent}) большая разница между сигналом и шумом"
                                      f" на следующих частотах: {','.join(diff_larger_3)}")
                frame = frames[file_type]
                df.rename(columns={0: 'frq', 1: 'max_s', 2: 'max_n'}, inplace=True)
                df = df.assign(min_s=df['max_s'], min_n=df['max_n'])
                if frame.empty:
                    frames[file_type] = df
                else:
                    try:
                        frame['max_s'] = np.where(frame['max_s'] < df['max_s'], df['max_s'], frame['max_s'])
                        frame['min_s'] = np.where(frame['min_s'] > df['min_s'], df['min_s'], frame['min_s'])
                        frame['max_n'] = np.where(frame['max_n'] < df['max_n'], df['max_n'], frame['max_n'])
                        frame['min_n'] = np.where(frame['min_n'] > df['min_n'], df['min_n'], frame['min_n'])
                        frames[file_type] = frame
                    except ValueError:
                        errors.append(f"В файле {file.name} ({file.parent}) отсутствуют некоторые частоты")
            if errors:
                self.logging.warning('Нашлись ошибки, программа остановлена')
                self.info_value.emit('УПС!', '\n'.join(errors))
                self.event.clear()
                self.event.wait()
                self.status.emit(f"Ошибка при генерации сплошного спектра в папке «{self.name_dir}»")
                os.chdir(self.default_path)
                self.status_finish.emit('gen_lf_pemi', str(self))
                time.sleep(0.5)  # Не удалять, не успевает отработать emit status_finish. Может потом
                self.window_check.close()
                return
            # Добавляем и вычитаем разброс значений, если значения минимума и максимума равны
            for frame in frames:
                mask = frames[frame]['max_s'].eq(frames[frame]['min_s'])
                frames[frame].loc[mask, 'max_s'] *= (1 + self.values_spread / 100)
                frames[frame].loc[mask, 'min_s'] *= (1 - self.values_spread / 100)
                mask = frames[frame]['max_n'].eq(frames[frame]['min_n'])
                frames[frame].loc[mask, 'max_n'] *= (1 + self.values_spread / 100)
                frames[frame].loc[mask, 'min_n'] *= (1 - self.values_spread / 100)
            for number in self.set_number:
                self.now_doc += 1
                self.line_doing.emit(f'Генерируем файл {number} ({self.now_doc} из {self.all_doc})')
                self.event.wait()
                if self.window_check.stop_threading:
                    raise CancelException()
                self.logging.info(f'Генерируем файл {number} ({self.now_doc} из {self.all_doc})')
                for frame in frames:
                    frq = frames[frame]['frq'].to_numpy().tolist()
                    max_s = frames[frame]['max_s'].to_numpy().tolist()
                    min_s = frames[frame]['min_s'].to_numpy().tolist()
                    sig = [random.uniform(maxs, mins) for maxs, mins in zip(max_s, min_s)]
                    max_n = frames[frame]['max_n'].to_numpy().tolist()
                    min_n = frames[frame]['min_n'].to_numpy().tolist()
                    noise = [random.uniform(maxn, minn) for maxn, minn in zip(max_n, min_n)]
                    # На всякий случай проверяем значения и меняем, если необходимо
                    for i in range(len(sig)):
                        if sig[i] < noise[i]:
                            sig[i], noise[i] = noise[i], sig[i]
                    folder = 'Эфир'
                    # Если нужно проверять на разницу, то гоняем, пока в каждой строке не окажется всё хорошо.
                    # Находим позиции, где не выполняется условие. Не изменяем шум, если он уже выше максимального
                    # и сигнал, если ниже минимального.
                    if frame in lines_file:
                        folder = 'Линии'
                        diff_3 = [True if s - n <= max_difference else False for s, n in zip(sig, noise)]
                        positions = {i: j for i, j in enumerate(diff_3) if j is False}
                        # Нужно брать разницу пополам и сразу прибавлять,
                        # если выходит за пределы - тогда только к одной стороне.
                        # По идее не должно быть такого, что обе разницы (сигнал и шум) сразу не удовлетворяют
                        # условию. Если будет так, то нужно будет редактировать процесс генерации.
                        for position in positions:
                            difference = (sig[position] - noise[position]
                                          - max_difference + random.uniform(0.1, 0.3))/2
                            if (((sig[position] - difference) > min_s[position])
                                    and ((noise[position] + difference) < max_n[position])):
                                diff = 0
                            elif (sig[position] - difference) < min_s[position]:
                                diff = difference - (sig[position] - min_s[position]) - random.uniform(0.1, 0.3)
                            else:
                                diff = -(difference - (max_n[position] - noise[position]) - random.uniform(0.1, 0.3))
                            sig[position] = sig[position] - difference + diff
                            noise[position] = noise[position] + difference + diff
                    writing_df = pd.DataFrame(data={'frq': frq, 'sig': sig, 'noise': noise}).round(2)
                    name_file = f'динамики.{frame}.txt' if frame != 'void' else 'динамики.txt'
                    path_df = Path(self.finish_path, str(number), folder, name_file)
                    path_df.parent.mkdir(parents=True, exist_ok=True)
                    writing_df.to_csv(path_df, header=False, index=False, sep='\t', mode='w', float_format="%.2f")
                current_progress += self.percent_progress
                self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                self.progress_value.emit(int(current_progress))
            self.progress_value.emit(100)
            self.status.emit(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» успешно завершена")
            self.logging.info(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» успешно завершена")
            os.chdir(self.default_path)
            self.status_finish.emit('gen_lf_pemi', str(self))
            time.sleep(0.5)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            # print(datetime.datetime.now() - start_time)
            return
        except CancelException:
            self.logging.warning(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» отменена пользователем")
            self.status.emit(f"Генерация НЧ ПЭМИ в папке «{self.name_dir}» отменена пользователем")
            os.chdir(self.default_path)
            self.status_finish.emit('gen_lf_pemi', str(self))
            time.sleep(0.5)  # Не удалять, не успевает отработать emit status_finish. Может потом
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
            time.sleep(0.5)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
