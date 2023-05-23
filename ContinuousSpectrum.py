import os
import pathlib
import random
import threading
import traceback

import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal
from convert import file_parcing


class GenerationFileCC(QThread):
    progress = pyqtSignal(int)  # Сигнал для progressbar
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
            df_sig = pd.DataFrame()
            df_noise = pd.DataFrame()
            df_list = [df_sig, df_noise]
            for folder in os.listdir(self.source):
                for file_csv in os.listdir(str(pathlib.Path(self.source, folder))):
                    self.logging.info('Читаем ' + file_csv)
                    if 'signal' in file_csv.lower():
                        names = ['frq', 'sig', 'str']
                        index_df = 0
                    else:
                        names = ['frq', 'noise', 'str']
                        index_df = 1
                    df = pd.read_csv(str(pathlib.Path(self.source, folder, file_csv)), delimiter=';',
                                     encoding="unicode_escape", header=None, names=names)
                    index = df[df['frq'] == 'Values'].index.to_list()[0]
                    df_write = df[0: index + 1]
                    df.drop(labels=[i for i in range(0, index + 1)], axis=0, inplace=True)
                    df.drop(labels=['str'], axis=1, inplace=True)
                    df = df.apply(pd.to_numeric, errors='coerce')
                    df.interpolate(inplace=True)
                    df_write = pd.concat([df_write, df])
                    self.logging.info('Перезапись файла ' + file_csv)
                    df_write.to_csv(str(pathlib.Path(self.source, folder, file_csv)), sep=';',
                                    header=False, index=False, encoding="ANSI")
                    df.set_index('frq', inplace=True)
                    df_list[index_df] = pd.concat([df_list[index_df], df])
                self.logging.info('Join и сортировка таблицы')
                all_data_df = df_list[0].join(df_list[1])
                all_data_df.sort_index(axis=0, inplace=True)
                path_dir = pathlib.Path(self.output, folder)
                os.makedirs(path_dir, exist_ok=True)
                path_dir = pathlib.Path(self.output, folder, folder + '.txt')
                path_dir.touch()
                # os.makedirs(path_dir, exist_ok=True)
                all_data_df.to_csv(path_dir, header=None, sep='\t', mode='w', float_format="%.8f")
            # index_sig = all_data_df[all_data_df['sig'] == 'nan'].index.to_list()[0]
            # index_noise = all_data_df[all_data_df['noise'] == 'nan'].index.to_list()[0]
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
