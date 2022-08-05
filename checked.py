import os
import re

import psutil


def checked_zone_checked(line_edit_path_check, line_edit_table_number, zone):
    def check(n, e):
        f = 0
        for el in e:
            if n == el:
                f = 1
                return f
        return f

    for proc in psutil.process_iter():
        if proc.name() == 'WINWORD.EXE':
            return ['УПС!', 'Закройте все файлы Word!']
    path = line_edit_path_check.text().strip()
    if not path:
        return ['УПС!', 'Путь к проверяемым документам пуст']
    if os.path.isdir(path):
        pass
    else:
        return ['УПС!', 'Указанный путь к проверяемым документам не является директорией']
    table = line_edit_table_number.text().strip()
    err_f = ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')
    if not table:
        return ['УПС!', 'Не указан номер таблицы']
    for i in table:
        flag = check(i, err_f)
        if not flag:
            return ['УПС!', 'Есть лишние символы в номере таблицы']
    zone_out = {i: el.text().replace(',', '.') for i, el in enumerate(zone) if el.text()}
    err_f = ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '.')
    err_msg = ('Есть лишние символы в ограничении по стационарным антеннам',
               'Есть лишние символы в ограничении по возимым антеннам',
               'Есть лишние символы в ограничении по носимым антеннам',
               'Есть лишние символы в ограничении r1',
               'Есть лишние символы в ограничении r1`')
    if zone_out:
        for i in zone_out:
            for j in zone_out[i]:
                flag = check(j, err_f)
                if not flag:
                    return ['УПС!', err_msg[i]]
    else:
        return ['УПС!', 'Не указано ни одно ограничение для проверки']
    return zone_out


def file_parcing_checked(dir_path, group_file):
    def folder_checked(p):
        errors = []
        txt_files = filter(lambda x: x.endswith('.txt'), os.listdir(p))
        if 'Описание.txt' not in txt_files:
            errors.append('Нет файла с описанием режимов (' + p + ')')
        else:
            with open(p + '\\Описание.txt', mode='r') as f:
                lines = f.readlines()
                for line in lines:
                    if re.findall(r'\s', line.rstrip('\n')):
                        errors.append('Пробелы в названии режимов (' + p + ', ' + line.rstrip('\n') + ')')
        return errors
    # Выбираем путь для исходников.
    path = dir_path.text().strip()
    if not path:
        return ['УПС!', 'Не указан путь к исходной папке']
    if os.path.isfile(path):
        return ['УПС!', 'Указанный путь к исходным файлам не является директорией']
    error = []
    if group_file:
        for folder in path:
            err = folder_checked(folder)
            if err:
                error.append(err)
    else:
        error = folder_checked(path)
    return ['УПС!', '\n'.join(error)] if error else {'path': path}
