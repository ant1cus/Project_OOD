import os
import re
import pathlib

import psutil
from openpyxl import load_workbook


def check(n, e):
    f = True
    for el in e:
        if n == el:
            f = False
            return f
    return f


def checked_zone_checked(line_edit_path_check, line_edit_table_number, zone):

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
    if not table:
        return ['УПС!', 'Не указан номер таблицы']
    for i in table:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')):
            return ['УПС!', 'Есть лишние символы в номере таблицы']
    zone_out = {i: el.text().replace(',', '.') for i, el in enumerate(zone) if el.text()}
    err_msg = ('Есть лишние символы в ограничении по стационарным антеннам',
               'Есть лишние символы в ограничении по возимым антеннам',
               'Есть лишние символы в ограничении по носимым антеннам',
               'Есть лишние символы в ограничении r1',
               'Есть лишние символы в ограничении r1`')
    if zone_out:
        for i in zone_out:
            for j in zone_out[i]:
                if check(j, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '.')):
                    return ['УПС!', err_msg[i]]
    else:
        return ['УПС!', 'Не указано ни одно ограничение для проверки']
    return zone_out


def file_parcing_checked(dir_path, group_file):
    def folder_checked(p):
        errors = []
        txt_files = filter(lambda x: x.endswith('.txt'), os.listdir(p))
        excel_files = [x for x in os.listdir(p) if x.endswith('.xlsx')]
        if 'Описание.txt' not in txt_files and 'описание.txt' not in txt_files:
            errors.append('Нет файла с описанием режимов (' + p + ')')
        else:
            with open(p + '\\Описание.txt', mode='r') as f:
                lines = f.readlines()
                for line in lines:
                    if re.findall(r'\s', line.rstrip('\n')):
                        errors.append('Пробелы в названии режимов (' + p + ', ' + line.rstrip('\n') + ')')
        return {'errors': errors, 'len': len(excel_files)}
    # Выбираем путь для исходников.
    path = dir_path.text().strip()
    if not path:
        return ['УПС!', 'Не указан путь к исходной папке']
    elif os.path.isfile(path):
        return ['УПС!', 'Указанный путь к исходным файлам не является директорией']
    else:
        folders = [i for i in os.listdir(path) if os.path.isdir(path + '\\' + i) and i != 'txt']
        if group_file is False and folders:
            return ['УПС!', 'В директории для парсинга присутствуют папки']
        elif group_file and folders is False:
            return ['УПС!', 'В директории для парсинга нет папок для преобразования']
    error = []
    progress = 0
    if group_file:
        for folder in os.listdir(path):
            err = folder_checked(path + '\\' + folder)
            progress += err['len']
            if err['errors']:
                error.append(err)
    else:
        err = folder_checked(path)
        error, progress = err['errors'], err['len']
    return ['УПС!', '\n'.join(error)] if error else {'path': path, 'progress': progress}


def check_generation_data(source_file, output_file, complect_number, complect_quantity, freq_restrict,
                          freq_restrict_path):

    source = source_file.text().strip()
    if not source:
        return ['УПС!', 'Путь к исходным файлам пуст']
    if os.path.isfile(source):
        return ['УПС!', 'Указанный путь к исходным файлам не является директорией']
    output = output_file.text().strip()
    if not output:
        return ['УПС!', 'Путь к создаваемым файлам пуст']
    if os.path.isfile(output):
        return ['УПС!', 'Указанный путь к создаваемым файлам не является директорией']
    if len([True for el in pathlib.Path(source).iterdir() if el.is_dir()]) > 1:
            return ['УПС!', 'В указанной директории слишком много папок']
    complect_num = complect_number.text().strip()
    if not complect_num:
        return ['УПС!', 'Не указаны номера комплектов']
    for i in complect_num:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', ' ', '-', ',', '.')):
            return ['УПС!', 'Есть лишние символы в номерах комплектов']
    complect_num = complect_num.replace(' ', '').replace(',', '.')
    if complect_num[0] == '.' or complect_num[0] == '-':
        return ['УПС!', 'Первый символ введён не верно']
    if complect_num[-1] == '.' or complect_num[-1] == '-':
        return ['УПС!', 'Последний символ введён не верно']
    for i in range(len(complect_num)):
        if complect_num[i] == '.' or complect_num[i] == '-':
            if complect_num[i + 1] == '.' or complect_num[i + 1] == '-':
                return ['УПС!', 'Два разделителя номеров подряд']
    complect = []
    for element in complect_num.split('.'):
        if '-' in element:
            num1, num2 = element.partition('-')[0], element.partition('-')[2]
            if num1 >= num2:
                return ['УПС!', 'Диапазон номеров комплектов указан не верно']
            else:
                for el in range(num1, num2 + 1):
                    complect.append(el)
        else:
            complect.append(element)
    complect.sort()
    if len(complect) != len(set(complect)):
        return ['УПС!', 'Есть повторения в номерах комплектов']
    complect_quant = complect_quantity.text()
    if not complect_quant:
        return ['УПС!', 'Не указано количество комплектов']
    for i in complect_quant:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')):
            return ['УПС!', 'Не правильно указано количество комплектов']
    if len(complect) != int(complect_quant):
        return ['УПС!', 'Указанные номера не совпадают с количеством генерируемых комплектов']
    txt_files = list(filter(lambda x: x.endswith('.txt'), os.listdir(source)))
    if len(txt_files) != 1:
        return ['УПС!', 'В папке больше одного или нет txt файла']
    name_mode = ''
    for file in sorted(txt_files):
        if os.stat(source + '\\' + file).st_size == 0:
            return ['УПС!', 'Файл с описанием режимов пуст']
        else:
            with open(source + '\\' + file, mode='r') as f:
                name_mode = f.readlines()
                if ' ' in name_mode:
                    return ['УПС!', 'Пробелы в названии режимов в файле «Описание»']
                else:
                    name_mode = [line.rstrip().lower() for line in name_mode if line]
    error = []
    for file_exel in sorted(list(filter(lambda x: x.endswith('.xlsx'), os.listdir(output)))):
        if file_exel[:-4] in complect:
            return ['УПС!', 'В указанной папке уже есть такие исходники']
        wb = load_workbook(source + '\\' + file_exel)  # Откроем книгу.
        name = wb.sheetnames  # Список листов.
        for name_exel in name:
            if name_exel.lower() not in name_mode:
                error.append('Названия режимов в файле ' + file_exel + ' не совпадают с описанием.')
                break
        wb.close()
    if error:
        return ['УПС!', '\n'.join(error)]
    # Добавление для разницы!
    restrict_file = False
    if freq_restrict:
        restrict_file = freq_restrict_path.text().strip()
        if not restrict_file:
            return ['УПС!', 'Путь к файлу с ограничениями пуст']
        if os.path.isfile(restrict_file):
            if restrict_file.endswith('.txt'):
                if os.stat(restrict_file).st_size == 0:
                    return ['УПС!', 'Файл для комплектов пуст']
            else:
                return ['УПС!', 'Файл для комплектов не txt']
        else:
            return ['УПС!', 'Указана директория в файле для ограничений']
    return {'source': source, 'output': output, 'complect': complect, 'complect_quant': complect_quant,
            'name_mode': name_mode, 'restrict_file': restrict_file}
