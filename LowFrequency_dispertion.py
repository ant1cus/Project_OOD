import os
import docx
import pandas as pd
import xlwings as xw
from decimal import Decimal
import threading
import traceback
import pythoncom
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from PyQt5.QtCore import QThread, pyqtSignal


class LFGeneration(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.path_old = incoming_data['source']
        self.path_new = incoming_data['output']
        self.excel = incoming_data['excel']
        self.logging = incoming_data['logging']
        self.q = incoming_data['q']
        self.event = threading.Event()

    def run(self):
        pythoncom.CoInitialize()
        current_progress = 0
        percent = 100 / len(os.listdir(self.path_old))
        self.logging.info('Начинаем генерацию НЧ')
        self.status.emit('Старт')
        self.progress.emit(current_progress)
        try:
            for element in os.listdir(self.path_old):
                self.status.emit('Вставляем данные в файл ' + element)
                self.logging.info('Вставляем данные в файл ' + element)
                self.logging.info('Открываем и закрываем excel для смены чисел')
                excel_app = xw.App(visible=False)
                excel_book = excel_app.books.open(self.excel)
                excel_book.save()
                excel_book.close()
                excel_app.quit()
                self.logging.info('Читаем таблицы')
                table_1 = pd.read_excel(self.excel, sheet_name='Таблица для вставки')
                table_2 = pd.read_excel(self.excel, sheet_name='Таблица для вставки с R2_r1_r1')
                self.logging.info('Открваем docx')
                doc = docx.Document(self.path_old + '\\' + element)
                table = doc.tables
                self.logging.info('Заполняем таблицу ' + str(len(table) - 3))
                for i, row in enumerate(table[len(table) - 3].rows):
                    if i != 0 and i != 1:
                        len_row = len(row.cells)
                        for j, cell in enumerate(row.cells):
                            self.logging.info(cell.text)
                            if j == 0:
                                name_device = cell.text.strip()
                                name_col = table_1.columns
                                data_row = table_1[table_1[name_col[1]] == name_device].index.to_list()[0]
                            elif 1 <= j < 4:
                                cell.text = str(Decimal(table_1.iloc[data_row, j + 1]).quantize(Decimal('1.0')))
                            elif j == 4:
                                cell.text = str(int(table_1.iloc[data_row, j + 1]))
                            elif j >= 5:
                                if len_row == 6:
                                    cell.text = str(Decimal(table_1.iloc[data_row, j + 2]).quantize(Decimal('1.0')))
                                    cell.paragraphs[
                                        0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Выравнивание по центру
                                    break
                                else:
                                    cell.text = str(Decimal(table_1.iloc[data_row, j + 1]).quantize(Decimal('1.0')))
                            cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Выравнивание по центру
                self.logging.info('Заполняем таблицу ' + str(len(table) - 2))
                for i, row in enumerate(table[len(table) - 2].rows):
                    data_row += 1
                    if i == 0:
                        flag = 1 if len(row.cells) == 2 else 0
                    else:
                        for j, cell in enumerate(row.cells):
                            if 'R2' not in cell.text and 'r1' not in cell.text and len(cell.text) != 0:
                                name_device = cell.text.strip()
                                name_col = table_2.columns
                                data_row = table_2[table_2[name_col[0]] == name_device].index.to_list()[0]
                            else:
                                if j > 0:
                                    if type(table_2.iloc[data_row, j + flag]) == str:
                                        cell.text = table_2.iloc[data_row, j + flag]
                                    else:
                                        cell.text = str(Decimal(table_2.iloc[data_row,
                                                                             j + flag]).quantize(Decimal('1.0')))
                                cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Выравнивание по центру
                self.logging.info('Сохраняем документ')
                doc.save(self.path_new + '\\' + element)
                current_progress += percent
                self.progress.emit(current_progress)
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
