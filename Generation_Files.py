import os
import pathlib
import random
import threading
import time
import traceback

import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal
from convert import file_parcing
from DoingWindow import CheckWindow


class CancelException(Exception):
    pass


class GenerationFile(QThread):
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
        self.set_quant = incoming_data['set_quant']
        self.name_mode = incoming_data['name_mode']
        self.restrict_file = incoming_data['restrict_file']
        self.no_freq_lim = incoming_data['no_freq_lim']
        self.no_excel_file = incoming_data['no_excel_file']
        self.db_diff = incoming_data['3db_difference']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()
        self.event.set()
        self.now_doc = 0
        self.all_doc = 0
        self.percent_progress = 0
        self.move = incoming_data['move']
        self.name_dir = pathlib.Path(self.output).name
        title = f'Генерация файлов ПЭМИ в папке «{self.name_dir}»'
        self.window_check = CheckWindow(self.default_path, self.event, self.move, title)
        self.progress_value.connect(self.window_check.progressBar.setValue)
        self.line_progress.connect(self.window_check.lineEdit_progress.setText)
        self.line_doing.connect(self.window_check.lineEdit_doing.setText)
        self.info_value.connect(self.window_check.info_message)
        self.window_check.show()

    def run(self):
        try:
            current_progress = 0
            self.logging.info('Начинаем генерировать файлы')
            self.all_doc = 2*(len(os.listdir(self.source)) - 1) + (1 if self.no_excel_file else 2)*int(self.set_quant)
            self.line_progress.emit(f'Выполнено {int(current_progress)} %')
            self.progress_value.emit(0)
            self.percent_progress = 100/self.all_doc
            self.line_doing.emit(f'Парсим файлы из исходной папки {pathlib.Path(self.source).name}')
            error = file_parcing(self.source, self.logging, self.line_doing, self.now_doc, self.all_doc,
                                 self.line_progress, self.progress_value, self.percent_progress, current_progress,
                                 self.no_freq_lim, self.default_path, self.event, self.window_check)
            if error['base_exception']:
                self.logging.error(error['text'])
                self.logging.error(error['trace'])
                self.logging.warning(f"Генерация файлов в папке «{self.name_dir}» не завершена из-за ошибки")
                self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
                self.event.clear()
                self.event.wait()
                self.status.emit(f"Ошибка при генерации файлов в папке «{self.name_dir}»")
                os.chdir(self.default_path)
                self.status_finish.emit('generate_pemi', str(self))
                time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
                self.window_check.close()
                return
            if error['cancel']:
                raise CancelException
            if error['error']:
                self.logging.info(f"В исходных файлах в папке «{pathlib.Path(self.source).name}» присутствуют ошибки")
                self.logging.info('\n'.join(error['error']))
                err = '\n' + '\n'.join(error['error'])
                self.info_value.emit('УПС!', f"Ошибки при парсинге файлов:{err}")
                self.event.clear()
                self.event.wait()
                self.status.emit(f"В исходных файлах в папке «{pathlib.Path(self.source).name}» присутствуют ошибки")
                os.chdir(self.default_path)
                self.status_finish.emit('generate_pemi', str(self))
                time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
                self.window_check.close()
                return
            self.now_doc = error['now_doc']
            quant_doc = len(os.listdir(self.source)) - 2
            current_progress = error['cp']
            self.line_progress.emit(f'Выполнено {int(current_progress)} %')
            self.progress_value.emit(int(current_progress))
            self.line_doing.emit(f'Считываем режимы из текстового файла')
            self.logging.info('Считываем режимы из текстового файла для создания словаря')
            self.event.wait()
            if self.window_check.stop_threading:
                raise CancelException()
            txt_files = filter(lambda x: x.endswith('.txt'), os.listdir(self.source))
            for file in sorted(txt_files):
                try:
                    with open(self.source + '\\' + file, mode='r', encoding="utf-8-sig") as f:
                        self.logging.info("Кодировка utf-8-sig")
                        mode_1 = f.readlines()
                        mode_1 = [line.rstrip() for line in mode_1]
                except UnicodeDecodeError:
                    with open(self.source + '\\' + file, mode='r') as f:
                        self.logging.info("Другая кодировка")
                        mode_1 = f.readlines()
                        mode_1 = [line.rstrip() for line in mode_1]
            mode = {x.lower(): pd.DataFrame() for x in mode_1 if x}
            description = pd.DataFrame()
            df_out = {x.lower(): pd.DataFrame(columns=['frq', 'max_s', 'min_s',
                                                       'max_n', 'min_n', 'quant_frq']) for x in mode_1}
            self.logging.info('Считываем значения из исходных файлов')
            for file in os.listdir(pathlib.Path(self.source, 'txt')):
                self.now_doc += 1
                self.line_doing.emit(f'Генерируем файл {file} ({self.now_doc} из {self.all_doc})')
                self.event.wait()
                if self.window_check.stop_threading:
                    raise CancelException()
                self.logging.info(f'Генерируем файл {file} ({self.now_doc} из {self.all_doc})')
                os.chdir(pathlib.Path(self.source, 'txt', file))
                for el in os.listdir():
                    if 'описание' not in el.lower():
                        if os.stat(r"./" + el).st_size != 0:
                            df = pd.read_csv(el, sep='\t', header=None)
                            if 1 in df.columns:
                                mode[el[:-4].lower()] = pd.concat([mode[el[:-4].lower()], df])
                            else:
                                mode[el[:-4].lower()] = pd.concat([mode[el[:-4].lower()],
                                                                   pd.Series(dtype='object')])
                        else:
                            mode[el[:-4].lower()] = pd.concat([mode[el[:-4].lower()], pd.Series(dtype='object')])
                    else:
                        if description.empty:
                            try:
                                description = pd.read_csv(el, sep='\t', header=None)
                            except (BaseException,):
                                description = pd.DataFrame()
                current_progress += self.percent_progress
                self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                self.progress_value.emit(int(current_progress))
            for file in mode:
                if 0 in mode[file].columns:
                    df = mode[file][0].value_counts()
                    for el in df.index.values:
                        col = (mode[file][mode[file][0].isin([el])].sort_values(by=[0]))
                        df_out[file] = pd.concat([df_out[file], pd.DataFrame({'frq': [el],
                                                                              'max_s': [col[1].max()],
                                                                              'min_s': [col[1].min()],
                                                                              'max_n': [col[2].max()],
                                                                              'min_n': [col[2].min()],
                                                                              'quant_frq': [df[el]/quant_doc]})],
                                                 axis=0)
            for set_number in self.set:
                self.now_doc += 1
                self.line_doing.emit(f'Генерируем файл {set_number} ({self.now_doc} из {self.all_doc})')
                self.event.wait()
                if self.window_check.stop_threading:
                    raise CancelException()
                self.logging.info(f'Генерируем файл {set_number} ({self.now_doc} из {self.all_doc})')
                df_sheet = {}
                for file in df_out:
                    df = pd.DataFrame(columns=['frq', 'signal', 'noise'])
                    for i, row in enumerate(df_out[file].sort_values(by='frq').itertuples(index=False)):
                        if random.random() > (1-row[5]):
                            if row[1] == row[2]:
                                s = random.uniform(row[1] + 1, row[2] - 1)
                            else:
                                s = random.uniform(row[1], row[2])
                            if row[3] == row[4]:
                                n = random.uniform(row[3] + 1, row[4] - 1)
                            else:
                                n = random.uniform(row[3], row[4])
                            if self.db_diff and (s-n) < 3:
                                while True:
                                    s = s + 0.1
                                    n = n - 0.1 if (n > row[4]) else n
                                    if (s-n) > 3.2:
                                        break
                            if self.no_freq_lim is False:
                                if s < n:
                                    s, n = n, s
                                if s == n:
                                    s = s + 0.5
                            if self.restrict_file:
                                df_limit = pd.read_csv(self.restrict_file, sep='\t', names=['Mode', 'Freq', 'Lim'])
                                df_limit = df_limit.replace({',': '.'}, regex=True)
                                for r in df_limit.itertuples(index=False):
                                    if file in r[0].lower():
                                        if float(r[1]) == row[0]:
                                            if (s - n) > float(r[2]):
                                                s = float(n) + float(r[2]) - random.uniform(0.01, 0.1)
                                            break
                            df = pd.concat([df, pd.DataFrame({'frq': [round(row[0], 4)], 'signal': [round(s, 2)],
                                                              'noise': [round(n, 2)]})], axis=0)
                    # df = df.round({'frq': 4, 'signal': 2, 'noise': 2})
                    # print(df.round({'frq': 4, 'signal': 2, 'noise': 2}))
                    df_sheet[file] = df

                # функция для создания txt или excel файла
                def create_file(data_for_create, path_for_create, sheet_for_create, wb_for_create):
                    if self.no_excel_file:
                        data_for_create.to_csv(path_for_create + sheet_for_create + '.txt',
                                               header=False, index=False, sep='\t', mode='a')
                    else:
                        data_for_create.to_excel(wb_for_create, sheet_name=sheet_for_create,
                                                 index=False, header=False)
                path_txt = self.output + '\\' + str(set_number) + '\\'
                wb = False
                if self.no_excel_file:
                    os.makedirs(self.output + '\\' + str(set_number))
                else:
                    wb = pd.ExcelWriter(str(pathlib.Path(self.output, str(set_number) + '.xlsx')))
                for sheet_name in df_sheet.keys():
                    if 'описание' in sheet_name.lower():
                        create_file(description, path_txt, sheet_name, wb)
                    elif df_sheet[sheet_name].empty:
                        data_for_write = pd.Series() if self.no_excel_file else pd.Series('Не обнаружено')
                        create_file(data_for_write, path_txt, sheet_name, wb)
                    else:
                        create_file(df_sheet[sheet_name], path_txt, sheet_name, wb)
                if self.no_excel_file is False:
                    wb.close()
                current_progress += self.percent_progress
                self.line_progress.emit(f'Выполнено {int(current_progress)} %')
                self.progress_value.emit(int(current_progress))
            self.logging.info('Создаём файл описания')
            if self.no_excel_file is False:
                with open(self.output + '\\Описание.txt', mode='w', encoding='utf-8-sig') as f:
                    f.write('\n'.join([el for el in mode]).rstrip())
                error = file_parcing(self.output, self.logging, self.line_doing, self.now_doc, self.all_doc,
                                     self.line_progress, self.progress_value, self.percent_progress, current_progress,
                                     self.no_freq_lim, self.default_path, self.event, self.window_check)
            if error['base_exception']:
                self.logging.error(error['text'])
                self.logging.error(error['trace'])
                self.logging.warning(f"Генерация файлов в папке «{self.name_dir}» не завершена из-за ошибки")
                self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
                self.event.clear()
                self.event.wait()
                self.status.emit(f"Ошибка при генерации файлов в папке «{self.name_dir}»")
                os.chdir(self.default_path)
                self.status_finish.emit('generate_pemi', str(self))
                time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
                self.window_check.close()
                return
            if error['cancel']:
                raise CancelException
            self.line_progress.emit(f'Выполнено 100 %')
            self.progress_value.emit(int(100))
            self.logging.info(f"Генрация файлов в папке «{self.name_dir}» успешно завершена")
            os.chdir(self.default_path)
            self.status.emit(f"Генрация файлов в папке «{self.name_dir}» успешно завершена")
            self.status_finish.emit('generate_pemi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            # print(datetime.datetime.now() - start_time)
            return
        except CancelException:
            self.logging.warning(f"Генерация файлов в папке «{self.name_dir}» отменена пользователем")
            self.status.emit(f"Генерация файлов в папке «{self.name_dir}» отменена пользователем")
            os.chdir(self.default_path)
            self.status_finish.emit('generate_pemi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.logging.warning(f"Генерация файлов в папке «{self.name_dir}» не завершена из-за ошибки")
            self.info_value.emit('УПС!', 'Работа программы завершена из-за непредвиденной ошибки')
            self.event.clear()
            self.event.wait()
            self.status.emit(f"Ошибка при генерации файлов в папке «{self.name_dir}»")
            os.chdir(self.default_path)
            self.status_finish.emit('generate_pemi', str(self))
            time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
            self.window_check.close()
            return
