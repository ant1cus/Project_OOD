import os
import threading

import docx
import openpyxl
import re
import traceback
from natsort import natsorted
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from PyQt5.QtCore import QThread, pyqtSignal


class ZoneChecked(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.path = incoming_data['path_check']
        self.table = incoming_data['table_number']
        self.department = incoming_data['department']
        self.win_lin = incoming_data['win_lin']
        self.zone = incoming_data['zone_all']
        self.one_table = incoming_data['one_table']
        self.logging = incoming_data['logging']
        self.q = incoming_data['q']
        self.event = threading.Event()

    def run(self):
        progress = 0
        self.logging.info('Начинаем проверять зоны')
        self.status.emit('Старт')
        self.progress.emit(progress)

        zone_name = ('Стац.', 'Воз.', 'Нос.', 'r1', 'r1`')
        os.chdir(self.path)
        self.logging.info("Сортировка")
        docs = [i for i in os.listdir('.') if i[-4:] == 'docx']
        docs = natsorted(docs)
        percent = 100 / len(docs)
        void = 0
        errors_for_excel = []
        self.logging.info("Проходимся по списку")
        if self.department:
            percent_ = percent / 10 if self.win_lin else percent / 5
        else:
            percent_ = percent / 4
        try:
            for name_doc in docs:
                if self.pause_threading():
                    return
                self.logging.info("Документ " + str(name_doc) + " в работе")
                self.status.emit('Проверяем документ ' + name_doc)
                if '~' not in name_doc:
                    # Для того, что бы не съезжала заливка ее нужно добавлять каждый раз для каждой ячейки
                    shading_elm = []
                    string = {}  # Для записи непроходящих частот
                    shading_index = 0  # Счетчик для заливки
                    doc = docx.Document(name_doc)
                    table = doc.tables[int(self.table) - 1]  # Таблица для проверки (общая)
                    if void == 1:
                        errors_for_excel.append('\n')
                        void = 2
                    errors = [name_doc.rpartition(' ')[2][:-5]]
                    win_lin = 10 if self.win_lin else 5  # Если 2 системы
                    zone = 0
                    self.logging.info("Считываем зоны")
                    if self.department:  # Если ФСБ
                        for j in range(0, win_lin, 1):
                            if self.pause_threading():
                                return
                            if j == 0 and self.win_lin:
                                errors.append('\n')
                                errors.append('Windows')
                                n_s = table.cell(2, 1).text  # Чтобы проверять и не считывать другую систему
                                name_system = re.findall(r'^\w+\b\s(\b\w+\b)', n_s)[0]
                            if j == 5 and self.win_lin:
                                if len(string) != 0:  # Добавляем частоты если они есть.
                                    errors.append(string)
                                    string = {}
                                errors.append('\n')
                                errors.append('Linux')
                                n_s = table.cell(6, 1).text  # Чтобы проверять имя системы
                                name_system = re.findall(r'^\w+\b\s(\b\w+\b\s\b\w+\b)', n_s)[0]
                            if j <= 2:
                                if self.win_lin:
                                    zone = table.cell(3, j + 2).text.replace(',', '.')
                                else:
                                    zone = table.cell(2, j + 2).text.replace(',', '.')
                            elif j <= 4:
                                if self.win_lin:
                                    zone = table.cell(j + 1, 2).text.replace(',', '.')
                                else:
                                    zone = table.cell(j, 2).text.replace(',', '.')
                            elif j <= 7 and self.win_lin:
                                zone = table.cell(7, j - 3).text.replace(',', '.')
                            elif j <= 9 and self.win_lin:
                                zone = table.cell(j, 2).text.replace(',', '.')

                            if void == 0:  # Добавляем имена диапазонов
                                errors_for_excel.append('\t')
                                for k in zone_name:
                                    errors_for_excel.append(k)
                                void = 2
                                errors_for_excel.append('\n')
                            void = 1
                            try:
                                errors.append(round(float(zone), 1))
                            except ValueError:
                                if '<' in zone:
                                    errors.append(zone)
                            if self.one_table is False and '<' not in zone:  # Если нужно проверять таблицу
                                self.logging.info("Проверяем и красим таблицу")
                                self.status.emit('Проверяем и закрашиваем таблицу в документе ' + str(name_doc))
                                if j in self.zone:
                                    # Условия проверки
                                    if (float(self.zone[j]) < float(zone)) and (float(self.zone[j]) != 0):
                                        flag_for_exit = 1  # Для прерывания цикла
                                        # Поисковый диапазон
                                        range_search = [9, 11, 13, 5, 7] * 2 if self.win_lin else [9, 11, 13, 5, 7]
                                        x = 0
                                        name = ''
                                        table_3 = doc.tables[1] if self.win_lin else doc.tables[2]
                                        for row in table_3.rows:
                                            for cell in row.cells:
                                                try:
                                                    # Для определения позиции ячейки (tc.top, tc.bottom etc)
                                                    tc = cell._tc
                                                    # Если полностью объединена
                                                    if tc.right - tc.left == len(table_3.columns):
                                                        if cell.text != 'Опасные сигналы не обнаружены':
                                                            if self.win_lin:
                                                                # Если находим имя - не прерываем цикл
                                                                if len(re.findall(name_system, cell.text)):
                                                                    flag_for_exit = 1
                                                                else:
                                                                    flag_for_exit = 0
                                                            # Имя системы
                                                            name = re.findall(r"\(([^)]*)\)", cell.text)[0]
                                                    if flag_for_exit:  # Если нужно в цикл
                                                        if tc.right == 1 and tc.left == 0:
                                                            frq = float(cell.text.replace(',', '.'))  # Частота
                                                            try:
                                                                x = float(table_3.cell(tc.top, range_search[j])
                                                                          .text.replace(',', '.'))
                                                            except BaseException:
                                                                if '<' in table_3.cell(tc.top, range_search[j]).text:
                                                                    x = -1
                                                            # Если больше, то красим через xml
                                                            if x > float(self.zone[j]):
                                                                string.setdefault(name, [])
                                                                shading_elm.append(parse_xml(
                                                                    r'<w:shd {} w:fill="FFFF00"/>'.format(nsdecls('w')))
                                                                )
                                                                table_3.rows[tc.top].cells[0]._tc.get_or_add_tcPr().append(
                                                                    shading_elm[shading_index])
                                                                shading_index += 1
                                                                shading_elm.append(parse_xml(
                                                                    r'<w:shd {} w:fill="FFFF00"/>'.format(nsdecls('w')))
                                                                )
                                                                table_3.rows[tc.top].cells[
                                                                    range_search[j]]._tc.get_or_add_tcPr().append(
                                                                    shading_elm[shading_index])
                                                                shading_index += 1
                                                                if frq not in string[name]:
                                                                    string[name].append(frq)
                                                    break
                                                except BaseException:
                                                    break
                                        self.logging.info("Сохраняем документ")
                                        doc.save(os.path.abspath(self.path) + '\\' + name_doc)
                            progress = progress + percent_
                            self.progress.emit(int(progress))
                    else:
                        self.logging.info("Считываем зоны")
                        for j in range(0, 4, 1):
                            if self.pause_threading():
                                return
                            zone = table.cell(j + 1, 1).text.replace(',', '.')

                            if void == 0:
                                errors_for_excel.append('\t')
                                for k in zone_name[:-1]:
                                    errors_for_excel.append(k)
                                void = 2
                                errors_for_excel.append('\n')
                            void = 1
                            try:
                                errors.append(int(zone))
                            except ValueError:
                                try:
                                    errors.append(round(float(zone), 1))
                                except ValueError:
                                    if '<' in zone:
                                        errors.append(zone)
                            if self.one_table is False:
                                self.logging.info("Проверяем и красим таблицу")
                                self.status.emit('Проверяем и закрашиваем таблицу в документе ' + str(name_doc))
                                if j in self.zone:
                                    # Условия проверки
                                    if (float(self.zone[j]) < float(zone)) and (float(self.zone[j]) != 0):
                                        flag_for_exit = 0
                                        range_search = [7, 8, 9, 10]
                                        x = 0
                                        table_3 = doc.tables[2]
                                        for row in table_3.rows:
                                            for cell in row.cells:
                                                try:
                                                    tc = cell._tc
                                                    if tc.right - tc.left == len(table_3.columns):
                                                        if cell.text == '3 категория':
                                                            flag_for_exit = 1
                                                        try:
                                                            try:
                                                                x = int(table_3.cell(tc.top, range_search[j]).text)
                                                            except ValueError:
                                                                x = float(table_3.cell(tc.top, range_search[j])
                                                                          .text.replace(',', '.'))
                                                        except BaseException:
                                                            if '<' in table_3.cell(tc.top, range_search[j]).text:
                                                                x = -1
                                                        if x > float(self.zone[j]):
                                                            # string.setdefault(name, [])
                                                            shading_elm.append(parse_xml(
                                                                r'<w:shd {} w:fill="FFFF00"/>'.format(nsdecls('w'))))
                                                            table_3.rows[tc.top].cells[1]._tc.get_or_add_tcPr().append(
                                                                shading_elm[shading_index])
                                                            shading_index += 1
                                                            shading_elm.append(parse_xml(
                                                                r'<w:shd {} w:fill="FFFF00"/>'.format(nsdecls('w'))))
                                                            table_3.rows[tc.top].cells[2]._tc.get_or_add_tcPr().append(
                                                                shading_elm[shading_index])
                                                            shading_index += 1
                                                            shading_elm.append(parse_xml(
                                                                r'<w:shd {} w:fill="FFFF00"/>'.format(nsdecls('w'))))
                                                            table_3.rows[tc.top].cells[
                                                                range_search[j]]._tc.get_or_add_tcPr().append(
                                                                shading_elm[shading_index])
                                                            shading_index += 1
                                                    break
                                                except BaseException:
                                                    break
                                            if flag_for_exit:
                                                break
                                        self.logging.info("Сохраняем документ")
                                        doc.save(os.path.abspath(self.path) + '\\' + name_doc)
                                progress = progress + percent_
                                self.progress.emit(int(progress))
                    self.logging.info("Добавляем результаты")
                    self.status.emit('Добавляем результаты документа ' + str(name_doc))
                    if void == 1:
                        for el in errors:
                            if type(el) != dict:
                                errors_for_excel.append(str(el))
                            else:
                                errors_for_excel.append(el)
                        if len(string) > 0:
                            errors_for_excel.append(string)
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            progress = 0
            self.progress.emit(progress)
            self.status.emit('Ошибка!')
            os.chdir('C://')
            try:
                if es.args[2][2] == 'Запрашиваемый номер семейства не существует.':
                    self.status.emit('Не верно указан номер таблицы')
            except IndexError:
                pass
            return
        try:
            if self.pause_threading():
                return
            errors_for_excel.append('\n')
            self.logging.info("Формируем excel")
            self.status.emit('Формируем отчёт')
            if self.department:
                zone = [self.zone.get(i) for i in range(0, 5)]
                len_fill = len(zone_name)
            else:
                zone = [self.zone.get(i) for i in range(0, 4)]
                len_fill = len(zone_name[:-1])
            thin = openpyxl.styles.Side(border_style="thin", color="000000")
            wb = openpyxl.Workbook()
            ws = wb.active
            i, j = 1, 1
            for el in errors_for_excel:
                if i != 1 and el != 'Windows' and el != 'Linux':  # Чтобы отдельно заполнять первый столбец
                    if el == '\n':  # Если новая строка (используется как разделитель)
                        pass
                    else:
                        if j != 1:  # Если столбец не первый
                            if '<' in el:  # Если значения '<0.1'
                                ws.cell(i, j).value = el  # Просто пишем
                            else:
                                if type(el) == dict:  # Если тип словарь (если не прошло и есть значения для записи)
                                    j_ = 0
                                    for key in el:  # Имя и значения. Форматирование.
                                        ws.cell(i, j + j_).value = key
                                        ws.cell(i, j + j_).alignment = openpyxl.styles.Alignment(horizontal="left",
                                                                                                 vertical="center")
                                        ws.cell(i, j + j_).font = openpyxl.styles.Font(bold=True,
                                                                                       name="Times New Roman",
                                                                                       size="11")
                                        j_ += 1
                                        for element in sorted(el[key]):
                                            ws.cell(i, j + j_).value = element
                                            ws.cell(i, j + j_).alignment = openpyxl.styles.Alignment(horizontal="left",
                                                                                                     vertical="center")
                                            ws.cell(i, j + j_).font = openpyxl.styles.Font(name="Times New Roman",
                                                                                           size="11")
                                            j_ += 1
                                elif '.' in el:  # Если есть точка, то форматируем
                                    ws.cell(i, j).number_format = '0.0'
                                    ws.cell(i, j).value = float(el)
                                else:  # Если целочисленные (для ФСТЭК в основном)
                                    ws.cell(i, j).value = int(el)
                        else:  # Если столбец первый
                            len_ = len(el.partition('.')[2])  # Определяем какая длина после запятой (номер комплекта)
                            ws.cell(i, j).number_format = '0.' + '0' * len_  # Для форматирования строки под требования
                            f = '.' + str(len_) + 'f'
                            ws.cell(i, j).value = float(format(float(el), f))  # Забиваем значение
                        if j < len_fill + 2:  # Форматируем ячейки
                            if j == 1:
                                ws.cell(i, j).alignment = openpyxl.styles.Alignment(horizontal="left",
                                                                                    vertical="center")
                            else:
                                ws.cell(i, j).border = openpyxl.styles.Border(top=thin, left=thin,
                                                                              right=thin, bottom=thin)
                                ws.cell(i, j).alignment = openpyxl.styles.Alignment(horizontal="center",
                                                                                    vertical="center")
                            ws.cell(i, j).font = openpyxl.styles.Font(name="Times New Roman", size="11")
                            ws.cell(i, j).fill = openpyxl.styles.PatternFill(start_color='92D050',
                                                                             end_color='92D050',
                                                                             fill_type="solid")
                else:  # Форматируем и вставляем заголовки.
                    ws.cell(i, j).value = el
                    if i != 1:
                        ws.cell(i, j).alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center")
                        ws.cell(i, j).fill = openpyxl.styles.PatternFill(start_color='92D050',
                                                                         end_color='92D050',
                                                                         fill_type="solid")
                    else:
                        ws.cell(i, j).alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
                    ws.cell(i, j).font = openpyxl.styles.Font(name="Times New Roman", size="11")
                j += 1  # Увеличиваем столбец
                if el == '\n':  # Если переход к новой строке.
                    flag = 0
                    for element in range(2, len_fill + 2):  # Проверяем на проходимость
                        if i != 1 and type(ws.cell(i, element).value) != str and ws.cell(i, element).value:
                            try:  # Если не проходит, то форматируем
                                if (ws.cell(i, element).value > float(zone[element - 2])) and\
                                        (float(zone[element - 2]) != 0):
                                    ws.cell(i, element).font = openpyxl.styles.Font(bold=True, name="Times New Roman")
                                    flag = 1
                            except TypeError:
                                pass
                    if flag:  # Если что-то форматируется, то закрашиваем
                        if self.win_lin:
                            for min_ in range(1, 3):  # Закрашиваем номер комплекта
                                if type(ws.cell(i - min_, 1).value) == float:
                                    ws.cell(i - min_, 1).alignment = openpyxl.styles.Alignment(
                                        horizontal="left", vertical="center")
                                    ws.cell(i - min_, 1).fill = openpyxl.styles.PatternFill(
                                        start_color='FFC000', end_color='FFC000', fill_type="solid")
                                    break
                        for element in range(1, len_fill + 2):  # Закрашиваем строки
                            ws.cell(i, element).fill = openpyxl.styles.PatternFill(
                                start_color='FFC000', end_color='FFC000', fill_type="solid")
                    i += 1
                    j = 1
            wb.save(filename='Зоны.xlsx')  # Сохраняем книгу.
            wb.close()
            os.startfile('Зоны.xlsx')
            progress = 100
            self.progress.emit(progress)
            self.status.emit('Готово!')
            os.chdir('C://')
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            progress = 0
            self.progress.emit(progress)
            self.status.emit('Ошибка!')
            os.chdir('C://')
            return

    def pause_threading(self):
        question = False if self.q.empty() else self.q.get_nowait()
        if question:
            self.messageChanged.emit('Вопрос?', 'Проверка остановлена пользователем. Нажмите «Да» для продолжения'
                                                ' или «Нет» для прерывания')
            self.event.wait()
            self.event.clear()
            if self.q.get_nowait():
                self.status.emit('Прервано пользователем')
                self.progress.emit(0)
                return True
        return False
