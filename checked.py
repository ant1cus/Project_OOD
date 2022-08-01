import os
from win32com.client import Dispatch
from PyQt5.QtWidgets import QMessageBox


def msgBox(title, text):
    msg = QMessageBox(QMessageBox.Critical, title, text)
    msg.exec_()


def checked_zone_checked(lineEdit_path_check, lineEdit_table_number, zone):
    def check(n, e):
        f = 0
        for el in e:
            if n == el:
                f = 1
                return f
        return f

    word = Dispatch("Word.Application")
    book = word.Documents.Count
    if book != 0:
        msgBox('УПС!', 'Закройте все файлы Word!')
        return
    path = lineEdit_path_check.text().strip()
    if not path:
        msgBox('УПС!', 'Путь к проверяемым документам пуст')
        return
    if os.path.isdir(path):
        pass
    else:
        msgBox('УПС!', 'Указанный путь к проверяемым документам не является директорией')
        return
    table = lineEdit_table_number.text().strip()
    err_f = ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')
    if not table:
        msgBox('УПС!', 'Не указан номер таблицы')
        return
    for i in table:
        flag = check(i, err_f)
        if not flag:
            msgBox('УПС!', 'Есть лишние символы в номере таблицы')
            return
    zone_out = {i: el.text().replace(',', '.') for i, el in enumerate(zone) if el.text()}
    err_f = ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '.')
    err_msg = ('Есть лишние символы в ограничении по стационарным антеннам',
               'Есть лишние символы в ограничении по возимым антеннам',
               'Есть лишние символы в ограничении по носимым антеннам',
               'Есть лишние символы в ограничении r1',
               'Есть лишние символы в ограничении r1`')
    print(zone_out)
    if zone_out:
        for i in zone_out:
            for j in zone_out[i]:
                flag = check(j, err_f)
                if not flag:
                    msgBox('УПС!', err_msg[i])
                    return
    else:
        msgBox('УПС!', 'Не указано ни одно ограничение для проверки')
        return
    return zone_out
