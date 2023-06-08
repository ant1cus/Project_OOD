import os
import pathlib
import itertools
import random
import threading
import traceback
import math

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
        self.percent_progress = 0

    def parcing(self, current_progress, path, generation, name, name_folder):
        def write_gen(df_gen, name_file, mode):
            for num_set in self.complect:
                self.logging.info('Запись генерируемых файлов ' + name_file + ' ' + str(num_set))
                self.status.emit('Запись генерируемых файлов ' + name_file + ' ' + str(num_set))
                path_dir_gen = pathlib.Path(self.output, str(num_set), name_file)
                try:
                    df_gen_old = pd.read_csv(path_dir_gen, delimiter=';',
                                             encoding="unicode_escape", header=None)
                except BaseException:
                    df_gen_old = pd.DataFrame()
                if len(df_gen.columns) == 3:
                    name_col = list(df_gen.columns)
                else:
                    name_col = list(df_gen.columns)
                    name_col.append('str')
                if len(list(df_gen)) == 3:
                    df_gen.columns = [0, 1, 2]
                else:
                    df_gen.columns = [0, 1]
                    df_gen[2] = ""
                df_to_write = df_gen
                if mode:
                    df_to_write[1] = df_to_write[1].apply(lambda x: random.uniform(x * 0.95, x * 1.05))
                df_write_new = pd.concat([df_gen_old, df_to_write])
                df_write_new.to_csv(path_dir_gen, sep=';',
                                    header=False, index=False, encoding="ANSI")
                df_gen.columns = name_col

        df_sig = pd.DataFrame()
        df_noise = pd.DataFrame()
        df_list = [df_sig, df_noise]
        for file_csv in [file for file in os.listdir(str(pathlib.Path(path)))
                         if file.lower().endswith('.csv')]:
            self.logging.info('Парсим ' + file_csv + ' для ' + name_folder)
            self.status.emit('Парсим ' + file_csv + ' для ' + name_folder)
            if 'sig' in file_csv.lower():
                names = ['frq', 'sig', 'str']
                index_df = 0
            else:
                names = ['frq', 'noise', 'str']
                index_df = 1
            df = pd.read_csv(str(pathlib.Path(path, file_csv)), delimiter=';',
                             encoding="unicode_escape", header=None, names=names)
            delimiter = False
            if df[names[1]].isna().sum() == len(df):  # проверка если разделитель не стандартный ";"
                df = pd.read_csv(str(pathlib.Path(path, file_csv)), delimiter=',',
                                 encoding="unicode_escape", header=None, names=names)
                delimiter = True
            index_val = df[df['frq'] == 'Values'].index.to_list()
            index_trace = df[df['frq'] == 'Trace'].index.to_list()
            if len(index_trace) > len(index_val):
                df = df.drop(labels=[i for i in range(index_trace[1], len(df))], axis=0)
            add_index = 6 if '180-' in file_csv.lower() else 0
            df_write = pd.DataFrame()
            df_to_concat = pd.DataFrame()
            for index, values in enumerate(index_val):
                index_start = 0 if index_val.index(values) == 0 else values - 3
                if delimiter:
                    delimiter = False
                    index_start += 1
                df_old = df[index_start: values + 1]
                if generation:
                    write_gen(df_old, file_csv, False)
                    self.logging.info('Парсим ' + file_csv + ' для ' + name_folder)
                    self.status.emit('Парсим ' + file_csv + ' для ' + name_folder)
                df_write = pd.concat([df_write, df_old])
                if len(index_val) != 1 and index + 1 < len(index_trace):
                    df_new = df.drop(labels=[i for i in range(index_trace[index + 1], len(df))], axis=0)
                else:
                    df_new = df
                df_new = df_new.drop(labels=[i for i in range(0, values + 1)], axis=0)
                df_new.drop(labels=['str'], axis=1, inplace=True)
                df_new = df_new.apply(pd.to_numeric, errors='coerce')
                df_new.interpolate(inplace=True)
                df_write = pd.concat([df_write, df_new])
                if generation:
                    write_gen(df_new, file_csv, True)
                    self.logging.info('Парсим ' + file_csv + ' для ' + name_folder)
                    self.status.emit('Парсим ' + file_csv + ' для ' + name_folder)
                    df_new.drop(labels=['str'], axis=1, inplace=True)
                df_new.set_index('frq', inplace=True)
                if len(index_val) != 1:
                    old_name_col = names[1]
                    new_name_col = names[1] + str(index + add_index)
                    df_new.rename(columns={old_name_col: new_name_col}, inplace=True)
                df_to_concat = df_new if df_to_concat.empty else df_to_concat.join(df_new)
            self.logging.info('Перезапись файла ' + file_csv)
            df_write.to_csv(str(pathlib.Path(path, file_csv)), sep=';',
                            header=False, index=False, encoding="ANSI")
            if add_index == 0:
                df_list[index_df] = pd.concat([df_list[index_df], df_to_concat])
            else:
                try:
                    df_list[index_df] = df_list[index_df].join(df_to_concat)
                except ValueError:
                    df_list[index_df].update(df_to_concat)
            current_progress += self.percent_progress
            self.progress.emit(int(current_progress))
        self.status.emit('Сортируем и записываем txt ')
        self.logging.info('Join и сортировка таблицы')
        all_data_df = df_list[0].join(df_list[1])
        all_data_df.sort_index(axis=0, inplace=True)
        if all_data_df.shape[1] > 2:
            col_name = [j + str(i) for i in range(0, 12) for j in ['sig', 'noise']]
            all_data_df = all_data_df[col_name]
        path_dir = pathlib.Path(path, name)
        path_dir.touch()
        self.logging.info('Запись текстового файла ' + name)
        all_data_df.to_csv(str(path_dir), header=None, sep='\t', mode='w', float_format="%.8f")
        return current_progress

    def run(self):
        try:
            progress = 0
            self.logging.info('Считывание файлов')
            self.status.emit('Старт')
            self.progress.emit(0)
            for file_csv in os.listdir(str(pathlib.Path(self.source))):
                progress += 1 if file_csv.lower().endswith('.csv') else 0
            self.logging.info('Создание папок для конечных файлов')
            quant_set = 0
            for number_complect in self.complect:
                quant_set += 1
                path_dir = pathlib.Path(self.output, str(number_complect))
                os.makedirs(path_dir, exist_ok=True)
            progress += progress*quant_set
            self.percent_progress = 100 / progress
            self.logging.info('Входные данные:')
            self.logging.info('0' + '-|-' + str(self.source) + '-|-' + 'True' + '-|-' + 'first.txt' + '-|-'
                              + str(self.source))
            current_progress = self.parcing(0, self.source, True, 'first.txt', 'исходного файла')
            for folder in os.listdir(str(pathlib.Path(self.output))):
                self.logging.info('Входные данные:')
                self.logging.info(str(current_progress) + '-|-' + str(pathlib.Path(self.output, str(folder))) + '-|-' +
                                  'False' + '-|-' + str(folder) + '.txt' + '-|-' + str(folder))
                current_progress = self.parcing(current_progress,
                                                str(pathlib.Path(self.output, str(folder))),
                                                False, str(folder) + '.txt', str(folder))

            # def write_gen(df_gen, name_file, mode):
            #     self.logging.info('Запись генерируемых файлов ' + name_file)
            #     self.status.emit('Запись генерируемых файлов ' + name_file)
            #     for num_set in self.complect:
            #         path_dir_gen = pathlib.Path(self.output, str(num_set), name_file)
            #         try:
            #             df_gen_old = pd.read_csv(path_dir_gen, delimiter=';',
            #                                      encoding="unicode_escape", header=None)
            #         except BaseException:
            #             df_gen_old = pd.DataFrame()
            #         name_col = list(df_gen.columns)
            #         if len(list(df_gen)) == 3:
            #             df_gen.columns = [0, 1, 2]
            #         else:
            #             df_gen.columns = [0, 1]
            #             df_gen[2] = ""
            #         df_to_write = df_gen
            #         if mode:
            #             df_to_write[1] = df_to_write[1].apply(lambda x: random.uniform(x * 0.95, x * 1.05))
            #         df_write_new = pd.concat([df_gen_old, df_to_write])
            #         df_write_new.to_csv(path_dir_gen, sep=';',
            #                             header=False, index=False, encoding="ANSI")
            #         df_gen.columns = name_col

            # for folder in os.listdir(self.source):
            #     for file_csv in [file for file in os.listdir(str(pathlib.Path(self.source, folder)))
            #                      if file.lower().endswith('.csv')]:
            #         self.logging.info('Читаем ' + file_csv)
            #         self.status.emit('Читаем ' + file_csv)
            #         if 'sig' in file_csv.lower():
            #             names = ['frq', 'sig', 'str']
            #             index_df = 0
            #         else:
            #             names = ['frq', 'noise', 'str']
            #             index_df = 1
            #         df = pd.read_csv(str(pathlib.Path(self.source, folder, file_csv)), delimiter=';',
            #                          encoding="unicode_escape", header=None, names=names)
            #         delimiter = False
            #         if df[names[1]].isna().sum() == len(df):  # проверка если разделитель не стандартный ";"
            #             df = pd.read_csv(str(pathlib.Path(self.source, folder, file_csv)), delimiter=',',
            #                              encoding="unicode_escape", header=None, names=names)
            #             delimiter = True
            #         index_val = df[df['frq'] == 'Values'].index.to_list()
            #         index_trace = df[df['frq'] == 'Trace'].index.to_list()
            #         if len(index_trace) > len(index_val):
            #             df = df.drop(labels=[i for i in range(index_trace[1], len(df))], axis=0)
            #         add_index = 6 if '180-' in file_csv.lower() else 0
            #         df_write = pd.DataFrame()
            #         df_to_concat = pd.DataFrame()
            #         for index, values in enumerate(index_val):
            #             index_start = 0 if index_val.index(values) == 0 else values - 3
            #             # index_stop = len(df) if len(index_val) == 1 else values - 3
            #             if delimiter:
            #                 delimiter = False
            #                 index_start += 1
            #             df_old = df[index_start: values + 1]
            #             write_gen(df_old, False)
            #             df_write = pd.concat([df_write, df_old])
            #             if len(index_val) != 1 and index + 1 < len(index_trace):
            #                 df_new = df.drop(labels=[i for i in range(index_trace[index + 1], len(df))], axis=0)
            #             else:
            #                 df_new = df
            #             df_new = df_new.drop(labels=[i for i in range(0, values + 1)], axis=0)
            #             df_new.drop(labels=['str'], axis=1, inplace=True)
            #             df_new = df_new.apply(pd.to_numeric, errors='coerce')
            #             df_new.interpolate(inplace=True)
            #             write_gen(df_new, True)
            #             df_write = pd.concat([df_write, df_new])
            #             df_new.set_index('frq', inplace=True)
            #             if len(index_val) != 1:
            #                 old_name_col = names[1]
            #                 new_name_col = names[1] + str(index + add_index)
            #                 df_new.rename(columns={old_name_col: new_name_col}, inplace=True)
            #             df_to_concat = df_new if df_to_concat.empty else df_to_concat.join(df_new)
            #         self.logging.info('Перезапись файла ' + file_csv)
            #         df_write.to_csv(str(pathlib.Path(self.source, folder, file_csv)), sep=';',
            #                         header=False, index=False, encoding="ANSI")
            #         # df_list[index_df].set_index('frq', inplace=True)
            #         if add_index == 0:
            #             df_list[index_df] = pd.concat([df_list[index_df], df_to_concat])
            #         else:
            #             try:
            #                 df_list[index_df] = df_list[index_df].join(df_to_concat)
            #             except ValueError:
            #                 df_list[index_df].update(df_to_concat)
            #         current_progress += percent_progress
            #         self.progress.emit(int(current_progress))
            #     self.status.emit('Сортируем и записываем txt ' + folder)
            #     self.logging.info('Join и сортировка таблицы')
            #     all_data_df = df_list[0].join(df_list[1])
            #     all_data_df.sort_index(axis=0, inplace=True)
            #     if all_data_df.shape[1] > 2:
            #         col_name = [j + str(i) for i in range(0, 12) for j in ['sig', 'noise']]
            #         all_data_df = all_data_df[col_name]
            #     path_dir = pathlib.Path(self.output, folder)
            #     os.makedirs(path_dir, exist_ok=True)
            #     path_dir = pathlib.Path(self.output, folder, folder + '.txt')
            #     path_dir.touch()
            #     # os.makedirs(path_dir, exist_ok=True)
            #     self.logging.info('Запись текстового файла')
            #     all_data_df.to_csv(str(path_dir), header=None, sep='\t', mode='w', float_format="%.8f")
            # for file_csv in [file for file in os.listdir(str(pathlib.Path(self.source)))
            #                  if file.lower().endswith('.csv')]:
            #     self.logging.info('Читаем ' + file_csv)
            #     self.status.emit('Читаем ' + file_csv)
            #     if 'sig' in file_csv.lower():
            #         names = ['frq', 'sig', 'str']
            #         index_df = 0
            #     else:
            #         names = ['frq', 'noise', 'str']
            #         index_df = 1
            #     df = pd.read_csv(str(pathlib.Path(self.source, file_csv)), delimiter=';',
            #                      encoding="unicode_escape", header=None, names=names)
            #     delimiter = False
            #     if df[names[1]].isna().sum() == len(df):  # проверка если разделитель не стандартный ";"
            #         df = pd.read_csv(str(pathlib.Path(self.source, file_csv)), delimiter=',',
            #                          encoding="unicode_escape", header=None, names=names)
            #         delimiter = True
            #     index_val = df[df['frq'] == 'Values'].index.to_list()
            #     index_trace = df[df['frq'] == 'Trace'].index.to_list()
            #     if len(index_trace) > len(index_val):
            #         df = df.drop(labels=[i for i in range(index_trace[1], len(df))], axis=0)
            #     add_index = 6 if '180-' in file_csv.lower() else 0
            #     df_write = pd.DataFrame()
            #     df_to_concat = pd.DataFrame()
            #     for index, values in enumerate(index_val):
            #         index_start = 0 if index_val.index(values) == 0 else values - 3
            #         # index_stop = len(df) if len(index_val) == 1 else values - 3
            #         if delimiter:
            #             delimiter = False
            #             index_start += 1
            #         df_old = df[index_start: values + 1]
            #         write_gen(df_old, file_csv, False)
            #         df_write = pd.concat([df_write, df_old])
            #         if len(index_val) != 1 and index + 1 < len(index_trace):
            #             df_new = df.drop(labels=[i for i in range(index_trace[index + 1], len(df))], axis=0)
            #         else:
            #             df_new = df
            #         df_new = df_new.drop(labels=[i for i in range(0, values + 1)], axis=0)
            #         df_new.drop(labels=['str'], axis=1, inplace=True)
            #         df_new = df_new.apply(pd.to_numeric, errors='coerce')
            #         df_new.interpolate(inplace=True)
            #         df_write = pd.concat([df_write, df_new])
            #         write_gen(df_new, file_csv, True)
            #         df_new.drop(labels=['str'], axis=1, inplace=True)
            #         df_new.set_index('frq', inplace=True)
            #         if len(index_val) != 1:
            #             old_name_col = names[1]
            #             new_name_col = names[1] + str(index + add_index)
            #             df_new.rename(columns={old_name_col: new_name_col}, inplace=True)
            #         df_to_concat = df_new if df_to_concat.empty else df_to_concat.join(df_new)
            #     self.logging.info('Перезапись файла ' + file_csv)
            #     df_write.to_csv(str(pathlib.Path(self.source, file_csv)), sep=';',
            #                     header=False, index=False, encoding="ANSI")
            #     # df_list[index_df].set_index('frq', inplace=True)
            #     if add_index == 0:
            #         df_list[index_df] = pd.concat([df_list[index_df], df_to_concat])
            #     else:
            #         try:
            #             df_list[index_df] = df_list[index_df].join(df_to_concat)
            #         except ValueError:
            #             df_list[index_df].update(df_to_concat)
            #     current_progress += percent_progress
            #     self.progress.emit(int(current_progress))
            # self.status.emit('Сортируем и записываем txt ')
            # self.logging.info('Join и сортировка таблицы')
            # all_data_df = df_list[0].join(df_list[1])
            # all_data_df.sort_index(axis=0, inplace=True)
            # if all_data_df.shape[1] > 2:
            #     col_name = [j + str(i) for i in range(0, 12) for j in ['sig', 'noise']]
            #     all_data_df = all_data_df[col_name]
            # path_dir = pathlib.Path(self.source, 'name.txt')
            # path_dir.touch()
            # # os.makedirs(path_dir, exist_ok=True)
            # self.logging.info('Запись текстового файла')
            # all_data_df.to_csv(str(path_dir), header=None, sep='\t', mode='w', float_format="%.8f")

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
