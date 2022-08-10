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


class GenerationFile(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)

    def __init__(self, output):  # Список переданных элементов.
        QThread.__init__(self)

        self.logging = output[6]
        self.q = output[7]
        self.event = threading.Event()

    def run(self):
        progress = 0
        self.logging.info("Начинаем")
        self.status.emit('Старт')
        self.progress.emit(progress)

def convert1(self):
    def createwb(p1, p3, name, df):

        os.chdir(p1)
        ind = 0
        j = 1
        wb = Workbook()
        spisok = n
        for index, row in df.iterrows():
            num = row['spisok']
            num_n = os.path.splitext(num)[0]
            os.chdir(p3)
            with open(r'./' + str(num), mode='r') as f:
                vals = []
                i = 1
                if ind == 0:
                    ind = 1
                    sheet = wb.active
                    wb.remove(sheet)
                else:
                    sheet = wb.create_sheet(str(num_n).lower())
                    for line in f:
                        j = 1
                        vals = line.split('\t')
                        indf = 0
                        for rec in vals:
                            # if len(rec) < 4:
                            #     pass
                            # else:
                            if indf == 0:
                                s = rec.rstrip()
                                c = sheet.cell(row=i, column=j)
                                c.value = float(s)
                                j = j + 1
                                indf = 1
                            else:
                                s = "%.2f" % float(rec)
                                c = sheet.cell(row=i, column=j)
                                c.value = float(s)
                                j = j + 1
                        i = i + 1
        os.chdir(p1)
        wb.save(r'./' + str(name) + '.xlsx')
        wb.close()

    def proverka(n, e):
        flag = 0
        for i in e:
            if n == i:
                flag = 1
                return flag
        return flag

    err = ''
    # Чтение файлов и проверка ошибок
    # Выбираем путь для исходников (отдельная папка).
    self.pbar.resetFormat()
    path1 = self.dirPath1.text()
    if not path1:
        QMessageBox.critical(self, 'УПС!', 'Путь к исходным файлам пуст')
        return
    if os.path.isdir(path1):
        print()
    else:
        QMessageBox.critical(self, 'УПС!', 'Указанный путь к исходным файлам не является директорией')
        return
    path2 = self.dirPath2.text()
    if not path2:
        QMessageBox.critical(self, 'УПС!', 'Путь к создаваемым файлам пуст')
        return
    if os.path.isdir(path2):
        print()
    else:
        QMessageBox.critical(self, 'УПС!', 'Указанный путь к создаваемым файлам не является директорией')
        return
    os.chdir(path2)
    spisok_v_pap2 = os.listdir()
    os.chdir(path1)
    spisok = os.listdir()
    spisokp = [1 for i in spisok if os.path.isdir(os.path.abspath(r'./' + i))]
    if len(spisokp) > 1:
        if self.checkBox1.isChecked():
            QMessageBox.critical(self, 'УПС!', 'В указанной директории слишком много папок')
            return
    namefile = self.nameEdit.text()
    if not namefile:
        QMessageBox.critical(self, 'УПС!', 'Не указаны номера комплектов')
        return
    err_f = ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', ' ', '-', ',', '.')
    for i in namefile:
        flag = proverka(i, err_f)
        if flag == 0:
            QMessageBox.critical(self, 'УПС!', 'Есть лишние символы в номерах комплектов')
            return
    namef = namefile.replace(' ', '')
    nfile = namefile.replace(' ', '')
    namef = namefile.replace(',', '.')
    nfile = namefile.replace(',', '.')
    nf = []
    err_f = ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')
    if namef[0] == '.':
        QMessageBox.critical(self, 'УПС!', 'Первый символ введён не верно')
        return
    if namef[0] == '-':
        QMessageBox.critical(self, 'УПС!', 'Первый символ введён не верно')
        return
    if namef[len(nfile) - 1] == '.':
        QMessageBox.critical(self, 'УПС!', 'Последний символ введён не верно')
        return
    if namef[len(nfile) - 1] == '-':
        QMessageBox.critical(self, 'УПС!', 'Последний символ введён не верно')
        return
    for i in range(len(namef)):
        if namef[i] == '.':
            if namef[i + 1] == '.':
                QMessageBox.critical(self, 'УПС!', 'Два разделителя номеров подряд')
                return
            if namef[i + 1] == '-':
                QMessageBox.critical(self, 'УПС!', 'Два разделителя номеров подряд')
                return
        if namef[i] == '-':
            flag = proverka(namef[i - 1], err_f)
            if flag == 0:
                QMessageBox.critical(self, 'УПС!', 'Не верно введён диапазон номеров')
                return
            flag = proverka(namef[i + 1], err_f)
            if flag == 0:
                QMessageBox.critical(self, 'УПС!', 'Не верно введён диапазон номеров')
                return
    number = self.quantityEdit.text()
    if not number:
        QMessageBox.critical(self, 'УПС!', 'Не указано количество комплектов')
        return
    for i in number:
        flag = proverka(i, err_f)
        if flag == 0:
            QMessageBox.critical(self, 'УПС!', 'Не правильно указано количество комплектов')
            return
    i = 0
    for index in namef:
        #      считаем элементы до запятой или тире, объединяем, записываем, сортируем, проверяем.
        if index == '.':
            if i == 0:
                nf.append(int(nfile.partition('.')[0]))
                nfile = nfile.partition('.')[2]
            else:
                i = 0
        if index == '-':
            i = 1
            a1 = int(nfile.partition('-')[0])
            a_m = nfile.partition('-')[2]
            a2 = int(a_m.partition('.')[0])
            if a1 >= a2:
                QMessageBox.critical(self, 'УПС!', 'Диапазон номеров комплектов указан не верно')
                return
            else:
                for numfile in range(a1, a2 + 1, 1):
                    nf.append(numfile)
                nfile = nfile.partition(',')[2]
        if not nfile:
            break
    if not nf:
        nf.append(int(nfile))
    else:
        if i == 0:
            if not nfile:
                print()
            else:
                nf.append(int(nfile))
    nf.sort()
    nf1 = set(nf)
    if len(nf) != len(nf1):
        QMessageBox.critical(self, 'УПС!', 'Есть повторения в номерах комплектов')
        return
    if len(nf) != int(number):
        QMessageBox.critical(self, 'УПС!', 'Указанные номера не совпадают с количеством генерируемых комплектов')
        return
    indfile = 0
    poz = 0
    file_op = os.path.abspath(r'./' + str(spisok[0]))
    txtfiles = list(filter(lambda x: x.endswith('.txt'), spisok))
    if len(txtfiles) != 1:
        QMessageBox.critical(self, 'УПС!', 'В папке больше одного или нет txt файла')
        return
    flag = 0
    it = []
    for file in sorted(txtfiles):
        if os.stat(r"./" + file).st_size == 0:
            QMessageBox.critical(self, 'УПС!', 'Файл с описанием режимов пуст')
            return
        else:
            path_txt_op = os.path.abspath(r'./' + str(file))
            with open(file, mode='r') as f:
                e = f.readlines()
            e = [line.rstrip() for line in e]
            e = [x for x in e if x]
            if self.checkBox1.isChecked():
                flag = 0
                error_1 = []
                exelfiles = list(filter(lambda x: x.endswith('.xlsx'), spisok_v_pap2))
                for file_exel in sorted(exelfiles):
                    for xl in nf:
                        if file_exel == (str(xl) + '.xlsx'):
                            QMessageBox.critical(self, 'УПС!', 'В указанной папке уже есть такие исходники')
                            return
                    wb = load_workbook("./" + file_exel)  # Откроем книгу.
                    name = wb.sheetnames  # Список листов.
                    i = 0
                    flag_err = 0
                    for line in e:
                        i = i + 1
                        for name_exel in name:
                            if line.lower() == name_exel.lower():
                                flag_err = flag_err + 1
                    wb.close()
                    if i != flag_err:
                        error_1.append('Названия режимов в файле ' + file_exel + ' не совпадают с описанием.\n')
                        flag = 1
            if flag == 1:
                err = ''.join(error_1)
                QMessageBox.critical(self, 'УПС!', err)
                return
            for line in e:
                try:
                    if line.index(' '):
                        f.close()
                        QMessageBox.critical(self, 'УПС!', 'Пробелы в названии режимов')
                        return
                except ValueError:
                    continue
    try:
        len(e)
    except UnboundLocalError:
        QMessageBox.critical(self, 'УПС!', 'Нет файла с описанием режимов')
        return

    # Новое добавление для разницы!
    if self.checkBox3.isChecked():
        path3 = self.dirPath3.text()
        if not path3:
            QMessageBox.critical(self, 'УПС!', 'Путь к файлу с ограничениями пуст')
            return
        if os.path.isfile(path3):
            if path3.endswith('.txt'):
                if os.stat(path3).st_size == 0:
                    QMessageBox.critical(self, 'УПС!', 'Файл для комплектов пуст')
                    return
            else:
                QMessageBox.critical(self, 'УПС!', 'Файл для комплектов не txt')
                return
        df_lim = pd.DataFrame({"Mode": [], "Freq": [], "Lim": []})
        with open(path3, mode='r') as f:
            df_lim = pd.read_csv(f, sep='\t', names=['Mode', 'Freq', 'Lim'])
            df_limit = df_lim.replace({',': '.'}, regex=True)

    df_sort = pd.DataFrame({"num": [], "spisok": []})
    si = 0
    i = 0
    txt = ''
    for file in e:
        df_sort.loc[si, 'num'] = i
        df_sort.loc[si, 'spisok'] = str(file) + '.txt'
        if i == 0:
            txt = str(file) + '.txt'
        i += 1
        si += 1
    i = 0
    # Если файлы exel.
    if self.checkBox1.isChecked():
        # Если есть exel преобразуем в txt.
        e = exel()
        err = e.exel(path1, indfile, poz)
        if len(err) != 0:
            shutil.rmtree(path1 + "/txt/")
            QMessageBox.critical(self, 'УПС!', err)
            return
        os.chdir(path1 + "/txt/")
        path1 = path1 + "/txt/"
    else:
        try:
            os.makedirs('./txt//')
        except (FileExistsError, AttributeError):
            print()
        path1 = path1 + "/txt/"
        spisok_pap = filter(lambda x: os.path.isdir(x), spisok)
        for el in spisok_pap:
            shutil.move(el, path1)
    os.chdir(path1)
    spisok_pap = os.listdir()
    flagdf = 0
    df = [0]
    if self.checkBox2.isChecked():
        prog = 100 / int(number)
    else:
        prog = 100 / (3 * int(number))
    percent = 0
    df_file = pd.DataFrame({"frq": [], "s": [], "n": []})
    # Метка для подсчёта.
    obpov = 1
    # Запись значений.
    for pap in spisok_pap:
        ind = 0
        os.chdir(r'./' + str(pap))
        df_sort['num'] = pd.to_numeric(df_sort['num'], errors='coerce')
        df_sort = df_sort.sort_values('num', ascending=True)
        for index, row in df_sort.iterrows():
            file = row['spisok']
            # Запись пути к описанию.
            if ind == 0:
                path_op = os.path.abspath(r'./' + str(txt))
            # Запись файлов.
            else:
                with open(r'./' + str(file), mode='r') as f:
                    if flagdf == 0:
                        # Частота, макс. сигнал, мин. сигнал, макс. шум, мин. шум,
                        # сколько раз появлялось в исходниках, общее кол-во исходников.
                        df.append(pd.DataFrame({
                            'frq': [0],
                            'maxs': [0],
                            'mins': [0],
                            'maxn': [0],
                            'minn': [0],
                            'fo': [0],
                            'f': [0]
                        }))
                        # Проверяем не является ли файл пустым.
                        if os.stat(r"./" + file).st_size == 0:
                            df[ind].loc[:, 'f'] = df[ind].loc[:, 'f'] + 1
                        else:

                            df_file = pd.read_csv(f, sep='\t', names=["frq", "s", "n"])
                            si = 0
                            for i in df_file.loc[:, 'frq']:
                                df[ind].loc[si, 'frq'] = df_file.loc[si, 'frq']
                                df[ind].loc[si, 'maxs':'mins'] = df_file.loc[si, 's']
                                df[ind].loc[si, 'maxn':'minn'] = df_file.loc[si, 'n']
                                df[ind].loc[si, 'fo'] = 1
                                df[ind].loc[si, 'f'] = obpov
                                si = si + 1
                    else:
                        # Проверяем не является ли файл пустым.
                        if os.stat(r"./" + file).st_size == 0:
                            df[ind].loc[:, 'f'] = obpov
                        df_file = pd.read_csv(f, sep='\t', names=["frq", "s", "n"])
                        si = 0
                        item = 0
                        # Сравниваем и записываем, если требуется, текущие значения в таблицы.
                        for i in df_file.loc[:, 'frq']:
                            item = 0
                            sj = 0
                            for j in df[ind].loc[:, 'frq']:
                                if df[ind].loc[sj, 'frq'] == df_file.loc[si, 'frq']:
                                    if df[ind].loc[sj, 'maxs'] < df_file.loc[si, 's']:
                                        df[ind].loc[sj, 'maxs'] = df_file.loc[si, 's']
                                    else:
                                        if df[ind].loc[sj, 'mins'] > df_file.loc[si, 's']:
                                            df[ind].loc[sj, 'mins'] = df_file.loc[si, 's']
                                    if df[ind].loc[sj, 'maxn'] < df_file.loc[si, 'n']:
                                        df[ind].loc[sj, 'maxn'] = df_file.loc[si, 'n']
                                    else:
                                        if df[ind].loc[sj, 'minn'] > df_file.loc[si, 'n']:
                                            df[ind].loc[sj, 'minn'] = df_file.loc[si, 'n']
                                    df[ind].loc[sj, 'fo'] = df[ind].loc[sj, 'fo'] + 1
                                    item = 1
                                    break
                                sj = sj + 1
                            if item == 0:
                                df[ind].loc[sj, 'frq'] = df_file.loc[si, 'frq']
                                df[ind].loc[sj, 'maxs':'mins'] = df_file.loc[si, 's']
                                df[ind].loc[sj, 'maxn':'minn'] = df_file.loc[si, 'n']
                                df[ind].loc[sj, 'fo'] = 1
                                df[ind].loc[sj, 'f'] = obpov
                            si = si + 1
                        df[ind].loc[:, 'f'] = obpov
            ind = ind + 1
        obpov = obpov + 1
        flagdf = 1
        os.chdir(path1)
    # Для каждой частоты разыгрываем random.choices с 0 и частотой, веса из появ./общ.появ.
    # Далее генерируем частоту и шум из диапазона (простой метод)
    df_s = pd.DataFrame({"frq": [], "s": [], "n": []})
    try:
        os.makedirs(path2 + "/txt/")
    except FileExistsError:
        print()
    path2 = path2 + "/txt/"
    os.chdir(path2)
    # Генерим файлы.
    for num in nf:
        ind = 0
        try:  # Изменение директории для работы.
            os.makedirs(path2 + str(num))
            os.chdir(path2 + str(num))
        except FileExistsError:
            continue
        # Создаём текстовые.
        for index, row in df_sort.iterrows():
            file = row['spisok']
            if ind == 0:
                shutil.copy2(str(path_op), str(path2 + str(num)))
            else:
                if re.findall(r'_lin', file) or re.findall(r'_linux', file):
                    file_name = str(file.swapcase()).lower()
                else:
                    file_name = str(file).lower()
                with open(r'./' + file_name, mode="w") as f:
                    si = 0
                    for freq in df[ind].loc[:, 'frq']:
                        pos = df[ind].loc[si, 'fo'] / df[ind].loc[si, 'f']
                        neg = (df[ind].loc[si, 'f'] - df[ind].loc[si, 'fo']) / df[ind].loc[si, 'f']
                        list_choice = [neg, pos]
                        frq = random.choices([0, freq], weights=list_choice)
                        if frq[0] == 0:
                            si = si + 1
                            continue
                        else:
                            frq_r = "%.4f" % float(frq[0])
                            if df[ind].loc[si, 'maxs'] == df[ind].loc[si, 'mins']:
                                s = random.uniform(df[ind].loc[si, 'maxs'] + 1, df[ind].loc[si, 'mins'] - 1)
                                s = "%.2f" % float(s)
                                n = random.uniform(df[ind].loc[si, 'maxn'] + 1, df[ind].loc[si, 'minn'] - 1)
                                n = "%.2f" % float(n)
                            else:
                                s = random.uniform(df[ind].loc[si, 'maxs'], df[ind].loc[si, 'mins'])
                                s = "%.2f" % float(s)
                                n = random.uniform(df[ind].loc[si, 'maxn'], df[ind].loc[si, 'minn'])
                                n = "%.2f" % float(n)
                            if self.checkBox3.isChecked():
                                j = 0
                                for i in df_limit.loc[:, 'Mode']:
                                    string = file.find(i)
                                    if string != -1:
                                        if float(df_limit.iloc[j, 1]) == float(frq_r):
                                            if (float(s) - float(n)) > float(df_limit.iloc[j, 2]):
                                                s = float(n) + float(df_limit.iloc[j, 2]) - random.uniform(0.01,
                                                                                                           0.1)
                                                s = "%.2f" % float(s)
                                            break
                                    j = j + 1
                            if float(s) < float(n):
                                s, n = n, s
                            if float(s) == float(n):
                                s = s + 0.1
                            print('{0:<}\t{1:>}\t{2:>}'.format(frq[0], s, n), file=f)
                        si = si + 1
                    # Сортировка.
                    with open(r'./' + file_name, mode='r') as f:
                        df_s = pd.read_csv(f, sep='\t', names=['frq', 's', 'n'])
                    df_s = df_s.sort_values(by=['frq'], ascending=True)
                    df_s = df_s.round({'frq': 4, 's': 2, 'n': 2})
                    df_s.reindex()
                    with open(r'./' + file_name, mode='w', newline='\n') as f:
                        df_s.to_csv(f, index=None, header=None, sep='\t')
            ind = ind + 1
        percent = percent + prog
        self.pbar.setValue(int(percent))
    if self.checkBox2.isChecked():
        percent = 100
        self.pbar.setValue(int(percent))
        self.pbar.setFormat('Готово!')
        return
    path2 = self.dirPath2.text()
    path_txt = path2 + "/txt/"
    flag = 0
    os.chdir(path1)
    spisok = os.listdir(path1)
    for file in spisok:
        if file.endswith(".xlsx"):
            flag = 1
    os.chdir(path_txt)
    spisok_pap = os.listdir()
    for pap in spisok_pap:
        for el in nf:
            if int(pap) == int(el):
                flagp = 0
                path3 = path_txt + str(pap)
                os.chdir(path3)
                if flag == 1:
                    os.chdir(path2)
                    exelfiles = filter(lambda x: x.endswith('.xlsx'), spisok)
                    for file in sorted(exelfiles):
                        if str(file) == str(pap + '.xlsx'):
                            flagp = 1
                            break
                    if flagp == 0:
                        createwb(path2, path3, str(pap), df_sort)
                else:
                    createwb(path2, path3, str(pap), df_sort)
                percent = percent + prog
                self.pbar.setValue(int(percent))
    os.chdir(path2)
    spisok = os.listdir(path2)
    i = 0
    exelfiles = list(filter(lambda x: x.endswith('.xlsx'), spisok))
    path1 = self.dirPath1.text()
    path2 = self.dirPath2.text()
    for file in sorted(exelfiles):
        for xl in nf:
            if file == (str(xl) + '.xlsx'):

                file_op2 = os.path.abspath(file)

                name1 = file_op.rpartition('\\')[2]
                name2 = file_op2.rpartition('\\')[2]
                xl = Dispatch("Excel.Application")
                if name1 == name2:
                    wb1 = xl.Workbooks.Open(Filename=(str(file_op)))
                    ws1 = wb1.Worksheets(1)
                    wb3 = xl.Workbooks.Add()
                    ws1.Copy(Before=wb3.Worksheets(1))
                    wb1.Close()
                    wb2 = xl.Workbooks.Open(Filename=(str(file_op2)))
                    ws1 = wb3.Worksheets(1)
                    ws1.Copy(Before=wb2.Worksheets(1))
                    wb3.Close(SaveChanges=False)
                else:
                    wb1 = xl.Workbooks.Open(Filename=(str(file_op)))
                    wb2 = xl.Workbooks.Open(Filename=(str(file_op2)))
                    ws1 = wb1.Worksheets(1)
                    ws1.Copy(Before=wb2.Worksheets(1))
                    wb1.Close()

                wb2.Close(SaveChanges=True)
                xl.Quit()
                percent = percent + prog
                self.pbar.setValue(percent)
    if path1 == path2:
        self.pbar.setFormat('Готово!')
    else:
        shutil.copy2(str(path_txt_op), str(path2))
    percent = 100
    self.pbar.setValue(int(percent))
    self.pbar.setFormat('Готово!')