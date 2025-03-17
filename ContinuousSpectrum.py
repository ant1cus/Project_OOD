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


class GenerationFileCC(QThread):
    status_finish = pyqtSignal(str, str)
    progress_value = pyqtSignal(int)
    info_value = pyqtSignal(str, str)
    status = pyqtSignal(str)
    line_progress = pyqtSignal(str)
    line_doing = pyqtSignal(str)

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.source = incoming_data['source']
        self.output = incoming_data['output']
        self.set = incoming_data['set']
        self.frequency = incoming_data['frequency']
        self.only_txt = incoming_data['only_txt']
        self.dispersion = incoming_data['dispersion']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()
        self.event.set()
        self.move = incoming_data['move']
        self.all_doc = 0
        self.now_doc = 0
        self.percent_progress = 0
        self.gen_txt = pd.DataFrame()
        self.name_txt = ''
        self.name_dir = pathlib.Path(self.source).name
        title = f'Генерация сплошного спектра в папке «{self.name_dir}»'
        self.window_check = CheckWindow(self.default_path, self.event, self.move, title)
        self.progress_value.connect(self.window_check.progressBar.setValue)
        self.line_progress.connect(self.window_check.lineEdit_progress.setText)
        self.line_doing.connect(self.window_check.lineEdit_doing.setText)
        self.info_value.connect(self.window_check.info_message)
        self.window_check.show()

    def parcing(self, current_progress, path, path_txt, generation, name_folder):
        def write_gen(df_gen, name_file, mode):
            for num_set in self.set:
                self.logging.info('Запись генерируемых файлов ' + name_file + ' ' + str(num_set))
                self.line_doing.emit(f'Запись генерируемых файлов {name_file} {str(num_set)}'
                                     f' ({self.now_doc} из {self.all_doc})')
                path_dir_gen = pathlib.Path(self.output, str(num_set), name_file)
                if os.path.exists(path_dir_gen):
                    df_gen_old = pd.read_csv(path_dir_gen, delimiter=';', encoding="unicode_escape", header=None)
                else:
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
                df_to_write = df_gen.copy()
                if mode:
                    df_to_write[1] = df_to_write[1].apply(lambda x: random.uniform(x * (1 - self.dispersion / 100),
                                                                                   x * (1 + self.dispersion / 100)))
                df_write_new = pd.concat([df_gen_old, df_to_write])
                df_write_new.to_csv(path_dir_gen, sep=';', header=False, index=False, encoding="ANSI")
                df_gen.columns = name_col

        df_sig = pd.DataFrame()
        df_noise = pd.DataFrame()
        df_list = [df_sig, df_noise]
        name = 'first'
        for file_csv in [file for file in os.listdir(str(pathlib.Path(path)))
                         if file.lower().endswith('.csv')]:
            self.event.wait()
            if self.window_check.stop_threading:
                raise CancelException()
            self.now_doc += 1
            if '_0-180_' in file_csv:
                name = file_csv.partition('_0-180_')[0] + '.txt'
            elif '_180-360_' in file_csv:
                name = file_csv.partition('_180-360_')[0] + '.txt'
            else:
                name = file_csv.rpartition('_')[0].rpartition('_')[0] + '.txt'
            if self.only_txt:
                self.name_txt = name
            self.logging.info('Парсим ' + file_csv + ' для ' + name_folder)
            self.line_doing.emit(f'Парсим {str(file_csv)} для комплекта '
                                 f'{name_folder} ({self.now_doc} из {self.all_doc})')
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
            if str(df.iloc[1, 1]) != '1.4':
                df.iloc[1, 1] = '1.4'
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
                if generation and self.set and self.only_txt is False:
                    write_gen(df_old, file_csv, False)
                    self.logging.info('Парсим ' + file_csv + ' для ' + name_folder)
                    self.line_doing.emit(f'Парсим {str(file_csv)} для комплекта '
                                         f'{name_folder} ({self.now_doc} из {self.all_doc})')
                df_write = pd.concat([df_write, df_old])
                if len(index_val) != 1 and index + 1 < len(index_trace):
                    df_new = df.drop(labels=[i for i in range(index_trace[index + 1], len(df))], axis=0)
                else:
                    df_new = df
                df_new = df_new.drop(labels=[i for i in range(0, values + 1)], axis=0)
                df_new.drop(labels=['str'], axis=1, inplace=True)
                # Меняем цифровой разделитель, если нужно.
                df_new[names[0]] = df_new[names[0]].str.replace(',', '.')
                df_new[names[1]] = df_new[names[1]].str.replace(',', '.')
                df_new = df_new.apply(pd.to_numeric, errors='coerce')
                df_new.dropna(how='all', inplace=True)
                df_new.interpolate(inplace=True)
                df_write = pd.concat([df_write, df_new])
                if generation and self.set and self.only_txt is False:
                    write_gen(df_new, file_csv, True)
                    self.logging.info('Парсим ' + file_csv + ' для ' + name_folder)
                    self.line_doing.emit(f'Парсим {str(file_csv)} для комплекта '
                                         f'{name_folder} ({self.now_doc} из {self.all_doc})')
                    df_new.drop(labels=['str'], axis=1, inplace=True)
                df_new['frq'] = df_new['frq'].astype(float)
                df_new['frq'] = df_new['frq'].multiply(0.000001)
                if self.frequency:
                    df_new = df_new[df_new.frq < float(self.frequency)]
                df_new.set_index('frq', inplace=True)
                if len(index_val) != 1:
                    old_name_col = names[1]
                    new_name_col = names[1] + str(index + add_index)
                    df_new.rename(columns={old_name_col: new_name_col}, inplace=True)
                df_to_concat = df_new if df_to_concat.empty else df_to_concat.join(df_new)
            self.logging.info('Перезапись файла ' + file_csv)
            if str(df_write.iloc[0, 0]) == 'sep=':  # Если при считывании эта строка первая, то будет криво писать
                df_write.drop(labels=[0], axis=0, inplace=True)
            df_write.to_csv(pathlib.Path(path, file_csv), sep=';',
                            header=False, index=False, encoding="ANSI")
            if add_index == 0:
                df_list[index_df] = pd.concat([df_list[index_df], df_to_concat])
            else:
                try:
                    df_list[index_df] = df_list[index_df].join(df_to_concat)
                except ValueError:
                    df_list[index_df].update(df_to_concat)
            current_progress += self.percent_progress
            self.progress_value.emit(int(current_progress))
            self.line_progress.emit(f'Выполнено {int(current_progress)} %')
        self.line_doing.emit(f'Сортируем и записываем txt ({self.now_doc} из {self.all_doc})')
        self.logging.info('Join и сортировка таблицы')
        all_data_df = df_list[0].join(df_list[1])
        all_data_df.sort_index(axis=0, inplace=True)
        if all_data_df.shape[1] > 2:
            col_name = [j + str(i) for i in range(0, 12) for j in ['sig', 'noise']]
            all_data_df = all_data_df[col_name]
        path_dir = pathlib.Path(path_txt)
        path_dir.mkdir(parents=True, exist_ok=True)
        path_dir = pathlib.Path(path_txt, name)
        path_dir.touch()
        self.logging.info('Запись текстового файла ' + name)
        all_data_df.to_csv(str(path_dir), header=False, sep='\t', mode='w', float_format="%.8f")
        if self.only_txt:
            self.gen_txt = all_data_df
        return current_progress

    def run(self):
        self.progress_value.emit(0)
        try:
            progress = 0
            self.logging.info('Считывание файлов')
            for file_csv in os.listdir(str(pathlib.Path(self.source))):
                progress += 1 if file_csv.lower().endswith('.csv') else 0
            self.logging.info('Создание папок для конечных файлов')
            quantity_set = 0
            if self.set and self.only_txt is False:
                for number_set in self.set:
                    quantity_set += 1
                    path_dir = pathlib.Path(self.output, str(number_set))
                    os.makedirs(path_dir, exist_ok=True)
            elif self.set and self.only_txt:
                quantity_set += len(self.set)
            progress += progress * quantity_set
            self.all_doc = progress
            self.percent_progress = 100 / progress
            self.line_progress.emit(f'Выполнено {0} %')
            self.logging.info('Входные данные:')
            self.logging.info('0' + '"|"' + str(self.source) + '"|"' + 'True' + '"|"' +
                              str(pathlib.PurePath(self.source).name) + '.txt' + '"|"' + str(self.source))
            self.line_doing.emit(f'Парсинг исходного файла ({self.now_doc} из {self.all_doc})')
            current_progress = self.parcing(0, self.source, str(pathlib.Path(self.source, 'txt')),
                                            True, 'исходного файла')
            self.line_progress.emit(f'Выполнено {int(current_progress)} %')
            self.event.wait()
            if self.window_check.stop_threading:
                raise CancelException()
            if self.set and self.only_txt:
                self.logging.info('Генерация без excel')
                for folder in self.set:
                    self.logging.info('Запись текстового файла ' + str(folder))
                    gen_df = self.gen_txt.apply(lambda x: random.uniform(x * (1 - self.dispersion / 100),
                                                                         x * (1 + self.dispersion / 100)))
                    path_df = pathlib.Path(self.output, 'txt', str(folder))
                    path_df.mkdir(parents=True, exist_ok=True)
                    path_df = pathlib.Path(self.output, 'txt', str(folder), self.name_txt)
                    gen_df.to_csv(str(path_df), header=None, sep='\t', mode='w', float_format="%.8f")
            elif self.set:
                self.logging.info('Генерация c excel')
                for folder in os.listdir(str(pathlib.Path(self.output))):
                    self.event.wait()
                    if self.window_check.stop_threading:
                        raise CancelException()
                    if os.path.isdir(str(pathlib.Path(self.output, folder))) and 'txt' not in folder:
                        self.logging.info('Входные данные:')
                        self.logging.info(str(current_progress) + '"|"' + str(pathlib.Path(self.output, str(folder))) +
                                          '"|"' + 'False' + '"|"' + str(folder) + '.txt' + '"|"' + str(folder))
                        self.line_doing.emit(f'Генерация {str(folder)} ({self.now_doc} из {self.all_doc})')
                        current_progress = self.parcing(current_progress,
                                                        str(pathlib.Path(self.output, str(folder))),
                                                        str(pathlib.Path(self.output, 'txt', str(folder))),
                                                        False, str(folder))
            self.progress_value.emit(100)
            self.status.emit(f"Генерация сплошного спектра в папке «{self.name_dir}» успешно завершена")
            self.logging.info(f"Генерация сплошного спектра в папке «{self.name_dir}» успешно завершена")
            os.chdir(self.default_path)
            self.status_finish.emit('continuous_spectrum', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            # print(datetime.datetime.now() - start_time)
            return
        except CancelException:
            self.logging.warning(f"Генерация сплошного спектра в папке «{self.name_dir}» отменена пользователем")
            self.status.emit(f"Генерация сплошного спектра в папке «{self.name_dir}» отменена пользователем")
            os.chdir(self.default_path)
            self.status_finish.emit('continuous_spectrum', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.logging.warning(f"Генерация сплошного спектра в папке «{self.name_dir}» не заврешена из-за ошибки")
            self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
            self.event.clear()
            self.event.wait()
            self.status.emit(f"Ошибка при генерации сплошного спектра в папке «{self.name_dir}»")
            os.chdir(self.default_path)
            self.status_finish.emit('continuous_spectrum', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
