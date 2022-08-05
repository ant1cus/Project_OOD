import os
import threading
import pandas as pd

import docx
import openpyxl
import re
import traceback
from natsort import natsorted
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from PyQt5.QtCore import QThread, pyqtSignal
from openpyxl import load_workbook


class FileParcing(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)

    def __init__(self, output):  # Список переданных элементов.
        QThread.__init__(self)
        self.path = output[0]
        self.group_file = output[1]
        self.logging = output[2]
        self.q = output[3]
        self.event = threading.Event()

    def run(self):
        progress = 0
        self.logging.info("Начинаем")
        self.status.emit('Старт')
        self.progress.emit(progress)

        err = ''
        spisok = os.listdir(self.path)
        # Сохраним нужное нам описание режимов.
        txtfiles = filter(lambda x: x.endswith('.txt'), spisok)
        for file in sorted(txtfiles):
            try:
                with open(self.path + '\\' + file, mode='r', encoding="utf-8-sig") as f:
                    mode_1 = f.readlines()
                    mode_1 = [line.rstrip() for line in mode_1]
            except UnicodeDecodeError:
                with open(self.path + '\\' + file, mode='r') as f:
                    mode_1 = f.readlines()
                    mode_1 = [line.rstrip() for line in mode_1]
        mode = [x for x in mode_1 if x]
        # Сюда вставить проверку на уже существующую папку и вопрос что делать
        try:
            os.makedirs(self.path + '\\txt\\')
        except (FileExistsError, AttributeError):
            pass
        # Работа с исходниками.
        # Отсортируем нужные нам файлы xlsx.
        exelfiles = filter(lambda x: x.endswith('.xlsx') and ('~' not in x), spisok)
        for file in sorted(exelfiles):
            wb = load_workbook(self.path + '\\' + file, data_only=True)  # Откроем книгу.
            namebook = str(file.rsplit('.xlsx', maxsplit=1)[0])  # Определение названия exel.
            name = wb.sheetnames  # Список листов.
            pat = ['_ЦП', '.m', '.v']  # список ключевых слов для поиска в ЦП
            pat_rez = ['_ЦП', '.m', '.v']
            for name_list in name:
                if re.search(r'_ЦП', name_list) or re.search(r'\.m', name_list) or re.search(r'\.v', name_list):
                    for elem in range(0, len(name)):  # поиск и устранение неточностей в названиях вкладок ЦП
                        if re.search(r'_ЦП', name[elem]) or re.search(r'\.m', name[elem]) or \
                                re.search(r'\.v', name[elem]):  # проверяем интересующие нас названия
                            rez = []
                            x = name[elem]
                            for y in pat:  # прогоняем список
                                if y == '.v':
                                    replace = re.findall(r'.v\d', x)
                                    if replace:
                                        y = replace[0]
                                        pat_rez[2] = y
                                rez.append(1) if x.find(y) != -1 else rez.append(-1)  # добавляем заметки для
                                # ключевых слов
                                x = x.replace(y, '')  # оставляем только название режима
                            for i in range(0, 3):
                                x = x + pat_rez[i] if rez[i] == 1 else x  # добавляем необходимые ключевые слова
                            worksheet = wb[name[elem]]  # выбираем лист с именем
                            worksheet.title = x  # переименовываем лист
                    wb.save(filename=file)  # сохраняем книгу
                    wb.close()
                    break

            wb = load_workbook(self.path + '\\' + file, data_only=True)  # Откроем книгу.
            name = wb.sheetnames  # Список листов.
            if name != mode:  # проверяем названия на соответствия
                err = str(err) + 'Названия режимов в исходнике' + str(file) + ' не совпадают с описанием:\n'
                for name_isx in name:
                    if mode.count(name_isx) == 0:
                        err = str(err) + 'режим ' + str(name_isx) + '\n'
            try:  # Исправить ошибку если уже есть такой файл внутри
                os.makedirs(self.path + '\\txt\\' + namebook)
                os.chdir(self.path + "\\txt\\" + namebook)
            except FileExistsError:
                continue
                # Загоняем в txt.
            for sheet in name:
                if sheet.lower() == 'описание':
                    df = pd.read_excel(self.path + '\\' + file, sheet_name=sheet, header=None)
                else:
                    df = pd.read_excel(self.path + '\\' + file, sheet_name=sheet, header=[0, 1, 2])
                    # df = df.dropna()
                    for row in df.itertuples(index=False):
                        if row[1]:
                            print(1)
                        print(row[0], '\n', row[1], '\n', row[2])
                        # if row('frq'):
                        #     if chek(s) and n is None:
                        #         error = 'В исходнике ' + b + ' в режиме ' + rezh + ' на частоте ' + str(
                        #             frq) + ' есть значение сигнала, но нет шума!\n'
                        #         err = str(err) + str(error)
                        #     elif chek(n) and s is None:
                        #         error = 'В исходнике ' + b + ' в режиме ' + rezh + ' на частоте ' + str(
                        #             frq) + ' есть значение шума, но нет сигнала!\n'
                        #         err = str(err) + str(error)
                        #     elif chek(s) and chek(n):
                df.to_csv(self.path + '\\txt\\' + namebook + '\\' + sheet + '.txt',
                          index=None, sep='\t', mode='w', header=None)
            wb.close()

        self.status.emit('Готово')

    def write_to_txt(self, w, nm, m, b, p):

        # Функция для проверки аргументов.
        def chek(arg):
            try:
                someVar = float(arg)
                return True
            except (TypeError, ValueError):
                return False

        ind = 0
        err = ''
        for poz in m:
            rezh = m[ind]
            k = 0
            for sheet in w:
                newlist = nm[k]
                newlist = newlist.replace(' ', '_')
                if rezh.lower() == newlist.lower():
                    i = 1
                    j = 1
                    # Создаем txt в выбранной директории.
                    name = rezh.lower()
                    if re.findall(r'_lin', name) or re.findall(r'_linux', name):
                        name = name.upper()
                    else:
                        name = name.lower()
                    with open(r'./' + name + ".txt", mode="w") as f:  # +str(ind)+"_"
                        if ind == 0:
                            # Определяем весь диапазон
                            rowm = sheet.max_row
                            colm = sheet.max_column
                            element = 0
                            for el in sheet:
                                # Запись первой страницы, на которой заметки.
                                for i in range(1, rowm + 1, 1):
                                    for j in range(1, colm + 1, 1):
                                        element = sheet.cell(i, j).value
                                        if element:
                                            print(str(element), file=f)
                                break
                        else:
                            for el in sheet:
                                # Считываем значения из exel.
                                frq = sheet.cell(row=i, column=j).value
                                s = sheet.cell(row=i, column=j + 1).value
                                n = sheet.cell(row=i, column=j + 2).value
                                # Записываем значения в txt, проверяя записи.
                                # Записи 'с метра не обнаружено' и похожие не выводятся.
                                if chek(frq):
                                    if chek(s) and n is None:
                                        error = 'В исходнике ' + b + ' в режиме ' + rezh + ' на частоте ' + str(
                                            frq) + ' есть значение сигнала, но нет шума!\n'
                                        err = str(err) + str(error)
                                    elif chek(n) and s is None:
                                        error = 'В исходнике ' + b + ' в режиме ' + rezh + ' на частоте ' + str(
                                            frq) + ' есть значение шума, но нет сигнала!\n'
                                        err = str(err) + str(error)
                                    elif chek(s) and chek(n):
                                        frq = int(frq) if str(frq).find('.') == -1 else round(float(
                                            str(frq).replace(', ', '.')), 4)
                                        s = "%.2f" % float(str(s).replace(', ', '.'))
                                        n = "%.2f" % float(str(n).replace(', ', '.'))
                                        if float(s) < float(n):
                                            error = 'В исходнике ' + b + ' в режиме ' + rezh + ' на частоте ' + str(
                                                frq) + ' значения шума больше сигнала!\n'
                                            err = str(err) + str(error)
                                        elif float(s) == float(n):
                                            error = 'В исходнике ' + b + ' в режиме ' + rezh + ' на частоте ' + str(
                                                frq) + ' одинаковые значения сигнала и шума!\n'
                                            err = str(err) + str(error)
                                        elif float(s) > 100:
                                            error = 'В исходнике ' + b + ' в режиме ' + rezh + ' на частоте ' + str(
                                                frq) + ' слишком большое значение сигнала!\n'
                                            err = str(err) + str(error)
                                        else:
                                            print('{0:<}\t{1:>}\t{2:>}'.format(frq, s, n), file=f)
                                i = i + 1
                                j = 1
                    k = k + 1
                else:
                    k = k + 1
            ind = ind + 1
        return err