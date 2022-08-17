import os
import random
import threading

import docx
import openpyxl
import re
import traceback

import pandas as pd
import numpy as np
from natsort import natsorted
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from PyQt5.QtCore import QThread, pyqtSignal
from convert import file_parcing


class GenerationFile(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.source = incoming_data['source']
        self.output = incoming_data['output']
        self.complect = incoming_data['complect']
        self.complect_quant = incoming_data['complect_quant']
        self.name_mode = incoming_data['name_mode']
        self.restrict_file = incoming_data['restrict_file']
        self.logging = incoming_data['logging']
        self.q = incoming_data['q']
        self.event = threading.Event()

    def run(self):
        current_progress = 0
        self.logging.info("Начинаем")
        self.status.emit('Старт')
        self.progress.emit(current_progress)
        percent = 100/(2*(len(os.listdir(self.source)) - 1) + int(self.complect_quant))
        error = file_parcing(self.source, self.logging, self.status, self.progress, percent, current_progress)
        quant_doc = len(os.listdir(self.source)) - 2
        errors = False
        if error['error']:
            errors = error['error']
        else:
            current_progress = error['cp']
            self.status.emit('Считываем режимы из текстового файла')
            self.logging.info('Считываем режимы из текстового файла для создания словаря')
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
            mode = {x: pd.DataFrame() for x in mode_1 if x}
            description = pd.DataFrame()
            df_out = {x: pd.DataFrame(columns=['frq', 'max_s', 'min_s', 'max_n', 'min_n', 'quant_frq']) for x in mode_1}
            self.status.emit('Считываем значения из исходных файлов')
            self.logging.info('Считываем значения из исходных файлов')
            for element in os.listdir(self.source + '\\' + 'txt'):
                self.logging.info('Файл ' + element)
                os.chdir(self.source + '\\' + 'txt' + '\\' + element)
                for el in os.listdir():
                    if 'описание' not in el.lower():
                        if os.stat(r"./" + el).st_size != 0:
                            df = pd.read_csv(el, sep='\t', header=None)
                            if 1 in df.columns:
                                mode[el[:-4]] = mode[el[:-4]].append(df)
                            else:
                                mode[el[:-4]] = mode[el[:-4]].append(pd.Series(), ignore_index=True)
                        else:
                            mode[el[:-4]] = mode[el[:-4]].append(pd.Series(), ignore_index=True)
                    else:
                        if description.empty:
                            description = pd.read_csv(el, sep='\t', header=None)
                current_progress += percent
                self.progress.emit(current_progress)
            for element in mode:
                if 0 in mode[element].columns:
                    df = mode[element][0].value_counts()
                    for el in df.index.values:
                        col = (mode[element][mode[element][0].isin([el])].sort_values(by=[0]))
                        df_out[element] = pd.concat([df_out[element], pd.DataFrame({'frq': [el],
                                                                                    'max_s': [col[1].max()],
                                                                                    'min_s': [col[1].min()],
                                                                                    'max_n': [col[2].max()],
                                                                                    'min_n': [col[2].min()],
                                                                                    'quant_frq': [df[el]/quant_doc]})],
                                                    axis=0)
            for complect_number in self.complect:
                self.logging.info('Генерация файла ' + str(complect_number))
                self.status.emit('Генерация файла ' + str(complect_number))
                wb = pd.ExcelWriter(self.output + '\\' + str(complect_number) + '.xlsx')
                df_sheet = {}
                for element in df_out:
                    df = pd.DataFrame(columns=['frq', 'signal', 'noise'])
                    for i, row in enumerate(df_out[element].itertuples(index=False)):
                        if random.random() > (1-row[5]):
                            if row[1] == row[2]:
                                s = random.uniform(row[1] + 1, row[2] - 1)
                            else:
                                s = random.uniform(row[1], row[2])
                            if row[3] == row[4]:
                                n = random.uniform(row[3] + 1, row[4] - 1)
                            else:
                                n = random.uniform(row[3], row[4])
                            if self.restrict_file:
                                df_limit = pd.read_csv(self.restrict_file, sep='\t', names=['Mode', 'Freq', 'Lim'])
                                df_limit = df_limit.replace({',': '.'}, regex=True)
                                for r in df_limit.itertuples(index=False):
                                    if element in r[0]:
                                        if float(r[1]) == row[0]:
                                            if (s - n) > float(r[2]):
                                                s = float(n) + float(r[2]) - random.uniform(0.01, 0.1)
                                            break
                            if s < n:
                                s, n = n, s
                            if s == n:
                                s = s + 0.5
                            df = pd.concat([df, pd.DataFrame({'frq': [row[0]], 'signal': [s], 'noise': [n]})], axis=0)
                    df = df.round({'frq': 4, 'signal': 2, 'noise': 2})
                    df_sheet[element] = df
                for sheet_name in df_sheet.keys():
                    if 'описание' in sheet_name.lower():
                        description.to_excel(wb, sheet_name=sheet_name, index=False, header=False)
                    else:
                        if df_sheet[sheet_name].empty:
                            pd.Series('Не обнаружено').to_excel(wb, sheet_name=sheet_name, index=False, header=False)
                        else:
                            df_sheet[sheet_name].to_excel(wb, sheet_name=sheet_name, index=False, header=False)
                wb.save()
                current_progress += percent
                self.progress.emit(current_progress)
            self.logging.info('Создаём файл описания')
            self.status.emit('Создаём файл описания')
            with open(self.output + '\\Описание.txt', mode='w', encoding='utf-8-sig') as f:
                f.write('\n'.join([el for el in mode]).rstrip())
        if errors:
            self.logging.info("Выводим ошибки")
            self.q.put({'errors_gen': errors})
            self.errors.emit()
            self.status.emit('В файлах присутствуют ошибки')
            self.progress.emit(0)
        else:
            self.progress.emit(100)
            self.logging.info("Конец работы программы")
            self.status.emit('Готово')
        os.chdir('C:\\')
        return
